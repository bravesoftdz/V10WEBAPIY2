unit YRechercherDestinataire;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, ExtCtrls, Grids, Hctrls, ComCtrls, HSysMenu, HTB97, Hent1,
  Mask,Utob,HMsgBox;

type
  TFRechercherDestinataire = class(TFVierge)
    PageRecherche: TPageControl;
    OngletCollaborateur: TTabSheet;
    GroupBox1: TGroupBox;
    GrilleCollaborateur: THGrid;
    OngletAnnuaire: TTabSheet;
    GroupBox2: TGroupBox;
    GrilleAnnuaire: THGrid;
    OngletRepertoire: TTabSheet;
    GroupBox3: TGroupBox;
    GrilleRepertoire: THGrid;
    Panel1: TPanel;
    ListeDiffusion: TComboBox;
    BDestA: TButton;
    BDestCc: TButton;
    GrilleCc: THGrid;
    GrilleA: THGrid;
    LCode: THLabel;
    Code: THCritMaskEdit;
    LNomCollaborateur: THLabel;
    LLogin: THLabel;
    NomCollaborateur: THCritMaskEdit;
    Groupe: THCritMaskEdit;
    LGroupe: THLabel;
    LDossier: THLabel;
    LNomAnnuaire: THLabel;
    LVille: THLabel;
    LFonction: THLabel;
    Fonction: THCritMaskEdit;
    Ville: THCritMaskEdit;
    NomAnnuaire: THCritMaskEdit;
    Dossier: THCritMaskEdit;
    LNomRepertoire: THLabel;
    LPrenom: THLabel;
    LEmail: THLabel;
    LCommentaire: THLabel;
    Commentaire: THCritMaskEdit;
    Email: THCritMaskEdit;
    Prenom: THCritMaskEdit;
    NomRepertoire: THCritMaskEdit;
    BRechercher: TToolbarButton97;
    lAide: THLabel;
    FindDialog: TFindDialog;
    BSupprimerDestA: TButton;
    BSupprimerDestCc: TButton;
    BCritereCollaborateur: TToolbarButton97;
    BCritereAnnuaire: TToolbarButton97;
    BCritereRepertoire: TToolbarButton97;
    Login: THCritMaskEdit;
    OngletListeDistribution: TTabSheet;
    GrilleListeDistribution: THGrid;
    procedure FormShow(Sender: TObject);
    procedure BDestAClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BDestCcClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure ModifierCritereCollaborateur (Sender: TObject);
    procedure ModifierCritereAnnuaire (Sender: TObject);
    procedure ModifierCritereRepertoire (Sender: TObject);
    procedure FindDialogFind (Sender: TObject);
    procedure BSupprimerDestAClick(Sender: TObject);
    procedure BSupprimerDestCcClick(Sender: TObject);

   private
    TobCollaborateur     : Tob;
    TobAnnuaire          : Tob;
    TobRepertoire        : Tob;
    TobListeDistribution : Tob;
    TobA                 : Tob;
    TobCc                : Tob;
    TobDestA             : Tob;
    TobDestCc            : Tob;

    procedure InitialiserGrilleCollaborateur ();
    procedure InitialiserGrilleAnnuaire ();
    procedure InitialiserGrilleRepertoire ();
    procedure InitialiserGrilleListeDistribution ();
    procedure InitialiserDestGrilleSelection (GrilleSelection : THGrid; TobDest : Tob; Var UneTob : Tob);
    procedure CopierDestGrilleSelection (GrilleSelection,GrilleDonnee : THGrid; Var UneTob : Tob);
    procedure SupprimerDestGrilleSelection (GrilleSelection : THGrid; Var UneTob : Tob);
    procedure RechercherPremiereLettrePremierecolonneGrille (UneGrille : ThGrid; Lettre: String);
    procedure DonnerInfoDestinataire (CoordonneeDestinataire : string; var TypeDestinataire : String; var LibelleDestinataire : String);


   Public
    ModeAffichageLite   : Boolean;
    NumDossier          : String;
    //DestA               : String;
    //DestCC              : String;
    DestinataireExterne : boolean;
  end;

//------------------------------------
//--- Déclaration procedure globale
//------------------------------------
function RechercherDestinataire (NumDossier : String; TobDestA : Tob; TobDestCc : Tob; Externe : Boolean) : String;
function LancerDestinataire (NumDossier : String; TobDestA : Tob; Externe : Boolean) : Boolean;

implementation
{$R *.DFM}

//---------------------------------------------
//--- Nom   : RechercherDestinataire
//--- Objet :
//---------------------------------------------
function LancerDestinataire (NumDossier : String; TobDestA : Tob; Externe : Boolean) : Boolean;
var FRechercherDestinataire: TFRechercherDestinataire;
begin
 FRechercherDestinataire:= TFRechercherDestinataire.Create(Application);
 FRechercherDestinataire.NumDossier:=NumDossier;
 FRechercherDestinataire.TobDestA:=TobDestA;
 FRechercherDestinataire.TobDestCc:=nil; 
 FRechercherDestinataire.DestinataireExterne:=Externe;
 FRechercherDestinataire.ModeAffichageLite:=True;
 try
  Result:=(FRechercherDestinataire.ShowModal = mrOk);
//  ListeDestA:=FRechercherDestinataire.DestA;
 finally
  FRechercherDestinataire.Free;
 end;
end;

//---------------------------------------------
//--- Nom   : RechercherDestinataire
//--- Objet :
//---------------------------------------------
function RechercherDestinataire (NumDossier : String; TobDestA : Tob; TobDestCc : Tob; Externe : Boolean) : String;
var FRechercherDestinataire: TFRechercherDestinataire;
begin
 FRechercherDestinataire:= TFRechercherDestinataire.Create(Application);
 FRechercherDestinataire.NumDossier:=NumDossier;
 FRechercherDestinataire.TobDestA:=TobDestA;
 FRechercherDestinataire.TobDestCc:=TobDestCc;
 FRechercherDestinataire.DestinataireExterne:=Externe;
 FRechercherDestinataire.ModeAffichageLite:=False;
 try
  FRechercherDestinataire.ShowModal;
//  ListeDestA:=FRechercherDestinataire.DestA;
//  ListeDestCc:=FRechercherDestinataire.DestCc;
 finally
  FRechercherDestinataire.Free;
 end;
end;

//---------------------------------------------
//--- Nom   : FormShow
//--- Objet :
//---------------------------------------------
procedure TFRechercherDestinataire.FormShow(Sender: TObject);
begin
 inherited;
 //--- Initialisation apparence des onglets
 OngletAnnuaire.TabVisible:=DestinataireExterne;
 OngletRepertoire.TabVisible:=DestinataireExterne;
 OngletListeDistribution.TabVisible:=DestinataireExterne;

 //--- Initialisation des grilles
 if (TobCollaborateur=nil) then InitialiserGrilleCollaborateur;
 if (OngletAnnuaire.TabVisible) and (TobAnnuaire=nil) then
  begin
   if (ModeAffichageLite) then Dossier.Text:='' else Dossier.Text:=NumDossier;
   InitialiserGrilleAnnuaire;
  end;
 if (OngletRepertoire.TabVisible) and (TobRepertoire=nil) then InitialiserGrilleRepertoire;

 if (ModeAffichageLite) then
  begin
   GrilleCc.Visible:=False;
   BDestCc.Visible:=False;
   BSupprimerDestCc.Visible:=False;
   BDestA.Caption:='Liste ->';
   OngletListeDistribution.TabVisible:=False;
   InitialiserGrilleListeDistribution;
  end;

 if (OngletListeDistribution.TabVisible) then InitialiserGrilleListeDistribution;

 InitialiserDestGrilleSelection (GrilleA,TobDestA,TobA);
 InitialiserDestGrilleSelection (GrilleCc,TobDestCc,TobCc);
end;

//---------------------------------------------
//--- Nom   : FormClose
//--- Objet :
//---------------------------------------------
procedure TFRechercherDestinataire.FormClose(Sender: TObject;var Action: TCloseAction);
begin
 inherited;
 if (TobCollaborateur<>nil) then TobCollaborateur.free;
 if (TobAnnuaire<>nil) then TobAnnuaire.free;
 if (TobRepertoire<>nil) then TobRepertoire.free;
 if (TobListeDistribution<>nil) then TobListeDistribution.free;
 if (TobA<>nil) then TobA.free;
 if (TobCc<>nil) then TobCc.free;
end;

//---------------------------------------------
//--- Nom   : InitialiserGrilleCollaborateur
//--- Objet :
//---------------------------------------------
procedure TFRechercherDestinataire.InitialiserGrilleCollaborateur ();
var ChSql            : String;
    Largeur          : Integer;
begin
 //--- Création de la tob annuaire
 GrilleCollaborateur.VidePile (False);
 if (TobCollaborateur=nil) then
  TobCollaborateur:=TOB.Create('Collaborateur', Nil, -1);

 if (TobCollaborateur<>nil) then
  begin
   //--- Création de la requête
   ChSql:='SELECT US_UTILISATEUR, US_ABREGE, US_LIBELLE FROM UTILISAT WHERE US_LIBELLE<>""';

   if (Code.Text<>'') then
    ChSql:=ChSql+' AND US_UTILISATEUR LIKE "'+Code.Text+'%"';
   {+GHA - 12/2007 FQ 11906}
   if (NomCollaborateur.text<>'') then
    ChSql:=ChSQl+' AND US_LIBELLE LIKE "'+NomCollaborateur.text+'%"';

   if (Login.text<>'') then
    ChSql:=ChSQl+' AND US_ABREGE LIKE "'+Login.text+'%"';
   {-GHA - 12/2007 FQ 11906}

   ChSql:=ChSQl+' ORDER BY US_UTILISATEUR';

   TobCollaborateur.LoadDetailFromSQL(ChSql);
   TobCollaborateur.PutGridDetail(GrilleCollaborateur, False, False, 'US_UTILISATEUR;US_ABREGE;US_LIBELLE');
  end;

 //--- Définition de la présentation de la grille
 Largeur := GrilleCollaborateur.Width div 8;
 GrilleCollaborateur.ColWidths[0] := Largeur;   // code
 GrilleCollaborateur.ColWidths[1] := 2*Largeur; // login
 GrilleCollaborateur.ColWidths[2] := 4*Largeur; // libellé
 HMTrad.ResizeGridColumns(GrilleCollaborateur);
end;

//----------------------------------------
//--- Nom   : InitialiserGrilleAnnuaire
//--- Objet :
//----------------------------------------
procedure TFRechercherDestinataire.InitialiserGrilleAnnuaire ();
var ChSql   : String;
    Largeur : Integer;
begin
 //--- Creation de la tob Annuaire
 GrilleAnnuaire.VidePile (False);
 if (TobAnnuaire=nil) then
  TobAnnuaire := TOB.Create('Annuaire et interlocuteurs', Nil, -1);

 if (TobAnnuaire<>nil) then
  begin
   //--- Création de la requête
   ChSql:='SELECT ANN_NOM1||" "||ANN_NOM2 AS ANN_NOM, DOS_NODOSSIER, ANN_ALCP, ANN_ALVILLE, "" AS C_CIVILITE, "" AS C_NOM, "" AS C_FONCTION, ANN_EMAIL AS C_RVA';
   ChSql:=ChSQl+' FROM ANNUAIRE LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER WHERE ANN_EMAIL<>""';

   if (Dossier.Text<>'') then
    ChSql:=ChSql+' AND DOS_NODOSSIER LIKE "'+Dossier.Text+'%"';
   if (NomAnnuaire.text<>'') then
    ChSql:=ChSQl+' AND ANN_NOM1 LIKE "'+NomAnnuaire.text+'%"';
   if (Ville.text<>'') then
    ChSql:=ChSQl+' AND ANN_ALVILLE LIKE "'+Ville.text+'%"';

   ChSql:=ChSQl+' UNION ';
        //mcd 12/2005 ChSql:=ChSQl+'SELECT ANN_NOM1||" "||ANN_NOM2 AS ANN_NOM, DOS_NODOSSIER, ANN_ALCP, ANN_ALVILLE, ANI_CV, ANI_NOM, ANI_FONCTIONJ, ANI_EMAIL';
        //mcd 12/2005 ChSql:=ChSQl+' FROM ANNUAIRE LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER LEFT JOIN ANNUINTERLOC ON ANN_GUIDPER=ANI_GUIDPER';
        //mcd 12/2005 ChSql:=ChSQl+' WHERE ANI_EMAIL<>""';
   ChSql:=ChSQl+'SELECT ANN_NOM1||" "||ANN_NOM2 AS ANN_NOM, DOS_NODOSSIER, ANN_ALCP, ANN_ALVILLE, C_CIVILITE, C_NOM, C_FONCTION, C_RVA';
   ChSql:=ChSQl+' FROM ANNUAIRE LEFT JOIN DOSSIER ON ANN_GUIDPER=DOS_GUIDPER LEFT JOIN CONTACT ON ANN_GUIDPER=C_GUIDPER';
   ChSql:=ChSQl+' WHERE C_RVA<>""';

   if (Dossier.Text<>'') then
    ChSql:=ChSql+' AND DOS_NODOSSIER LIKE "'+Dossier.Text+'%"';
   if (NomAnnuaire.text<>'') then
    ChSql:=ChSQl+' AND ANN_NOM1 LIKE "'+NomAnnuaire.text+'%"';
   if (Ville.text<>'') then
    ChSql:=ChSQl+' AND ANN_ALVILLE LIKE "'+Ville.text+'%"';

   ChSql:=ChSQl+' ORDER BY ANN_NOM';

   TobAnnuaire.LoadDetailFromSQL(ChSql);
      //mcd 12/2005 TobAnnuaire.PutGridDetail(GrilleAnnuaire,False,False,'ANN_NOM;DOS_NODOSSIER;ANN_ALCP;ANN_ALVILLE;ANI_CV;ANI_NOM;ANI_FONCTIONJ;ANI_EMAIL');
   TobAnnuaire.PutGridDetail(GrilleAnnuaire,False,False,'ANN_NOM;DOS_NODOSSIER;ANN_ALCP;ANN_ALVILLE;C_CIVILITE;C_NOM;C_FONCTION;C_RVA');
  end;

 //--- Définition de la présentation de la grille
 Largeur:=GrilleAnnuaire.Width div 19;
 GrilleAnnuaire.ColWidths[0]:=3*largeur;   // Nom1
 GrilleAnnuaire.ColWidths[1]:=2*largeur;   // NoDossier
 GrilleAnnuaire.ColWidths[2]:=2*largeur;   // code postal
 GrilleAnnuaire.ColWidths[3]:=2*largeur;   // ville
 GrilleAnnuaire.ColWidths[4]:=2*largeur;   // civilité
 GrilleAnnuaire.ColWidths[5]:=2*largeur;   // interlocuteur
 GrilleAnnuaire.ColWidths[6]:=2*largeur;   // fonction
 GrilleAnnuaire.ColWidths[7]:=4*largeur;   // email
 HMTrad.ResizeGridColumns(GrilleAnnuaire);
end;

//------------------------------------------
//--- Nom   : InitialiserGrilleRepertoire
//--- Objet :
//------------------------------------------
procedure TFRechercherDestinataire.InitialiserGrilleRepertoire();
var ChSql         : String;
    Largeur       : Integer;
begin
 //--- Creation de la tob Repertoire
 GrilleRepertoire.VidePile (False);
 if (TobRepertoire=nil) then
  TobRepertoire:=TOB.Create('Répertoire personnel', Nil, -1);

 if (TobRepertoire<>nil) then
  begin
   //--- Création de la requête
   ChSql:='SELECT YRP_NOM, YRP_PRENOM, YRP_CIVILITE, YRP_EMAIL FROM YREPERTPERSO WHERE YRP_EMAIL<>""';
   if (NomRepertoire.text<>'') then
    ChSql:=ChSQl+' AND YRP_NOM LIKE "'+NomRepertoire.text+'%"';
   if (Prenom.text<>'') then
    ChSql:=ChSQl+' AND YRP_PRENOM LIKE "'+Prenom.text+'%"';
   if (Email.text<>'') then
    ChSql:=ChSQl+' AND YRP_EMAIL LIKE "'+Email.text+'%"';

   ChSql:=ChSQl+' AND YRP_USER="'+V_PGI.User+'"';
   ChSql:=ChSql+' ORDER BY YRP_NOM';

   TobRepertoire.LoadDetailFromSQL(ChSql);
   TobRepertoire.PutGridDetail(GrilleRepertoire, False, False, 'YRP_NOM;YRP_PRENOM;YRP_CIVILITE;YRP_EMAIL');
  end;

 //--- Définition de la présentation de la grille
 Largeur:=GrilleRepertoire.Width div 14;
 GrilleRepertoire.ColWidths[0]:=3*Largeur;  // Nom
 GrilleRepertoire.ColWidths[1]:=3*Largeur;  // prénom
 GrilleRepertoire.ColWidths[2]:=2*Largeur;  // civilité
 GrilleRepertoire.ColWidths[3]:=6*Largeur;  // email
 HMTrad.ResizeGridColumns(GrilleRepertoire);
end;

//-------------------------------------------------
//--- Nom   : InitialiserGrilleListeDistribution
//--- Objet :
//-------------------------------------------------
procedure TFRechercherDestinataire.InitialiserGrilleListeDistribution ();
var ChSql            : String;
    Largeur,Indice   : Integer;
begin
 //--- Création de la tob annuaire
 GrilleListeDistribution.VidePile (False);
 if (TobListeDistribution=nil) then
  TobListeDistribution:=TOB.Create('ListeDistribution', Nil, -1);

 if (TobListeDistribution<>nil) then
  begin
   //--- Création de la requête
   ChSql:='SELECT YLI_LISTEDISTRIB, YLI_LIBELLELISTE, YLI_USER FROM YLISTEDISTRIB ORDER BY YLI_LISTEDISTRIB';
   TobListeDistribution.LoadDetailFromSQL(ChSql);

   for Indice:=0 to TobListeDistribution.Detail.Count-1 do
    begin
     if (TobListeDistribution.Detail [Indice].GetValue ('YLI_USER')<>'') then
      TobListeDistribution.Detail [Indice].AddChampSupValeur ('PRIVE','#ICO#3')
     else
      TobListeDistribution.Detail [Indice].AddChampSupValeur ('PRIVE','');
    end;

   TobListeDistribution.PutGridDetail(GrilleListeDistribution, False, False, 'YLI_LISTEDISTRIB;YLI_LIBELLELISTE;PRIVE');
  end;

 //--- Définition de la présentation de la grille
 Largeur := GrilleListeDistribution.Width div 8;
 GrilleListeDistribution.ColWidths[0] := 3*Largeur;  // Nom de la liste
 GrilleListeDistribution.ColWidths[1] := 4*Largeur;  // Description
 GrilleListeDistribution.ColWidths[2] := 2*Largeur;  // Privé
 GrilleListeDistribution.ColAligns[2] := TaCenter;
 HMTrad.ResizeGridColumns(GrilleListeDistribution);
end;

//-------------------------------------------
//--- Nom   : ModifierCritereCollaborateur
//--- Objet :
//-------------------------------------------
procedure TFRechercherDestinataire.ModifierCritereCollaborateur (Sender: TObject);
begin
 inherited;
 InitialiserGrilleCollaborateur;
end;

//--------------------------------------
//--- Nom   : ModifierCritereAnnuaire
//--- Objet :
//--------------------------------------
procedure TFRechercherDestinataire.ModifierCritereAnnuaire (Sender: TObject);
begin
 inherited;
 if not ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_NODOSSIER="'+Dossier.Text+'"') and (Dossier.Text<>'') then
  begin
   PgiInfo ('Le numéro de dossier est incohérent.',TitreHalley);
   Dossier.SetFocus;
  end
 else
  InitialiserGrilleAnnuaire;
end;

//-------------------------------------------
//--- Nom   : ModifierCritereRepertoire
//--- Objet :
//-------------------------------------------
procedure TFRechercherDestinataire.ModifierCritereRepertoire (Sender: TObject);
begin
 inherited;
 InitialiserGrilleRepertoire;
end;

//---------------------------------------------
//--- Nom   : InitialiserDestGrilleSelection
//--- Objet :
//---------------------------------------------
procedure TFRechercherDestinataire.InitialiserDestGrilleSelection (GrilleSelection : THGrid; TobDest : Tob; Var UneTob : Tob);
var TypeDestinataire       : String;
    CoordonneeDestinataire : String;
    LibelleDestinataire    : String;
    Largeur                : Integer;
    UneTobEnreg            : Tob;
    Indice                 : Integer;
begin
 //--- Creation de la tob
 if UneTob=nil then
  UneTob:=TOB.Create('', Nil, -1);

 if (UneTob<>nil) then
  begin
   GrilleSelection.VidePile (False);

   if (TobDest<>nil) then
    begin
     for Indice:=0 to TobDest.Detail.Count-1 do
      begin
       CoordonneeDestinataire:=TobDest.Detail [Indice].GetValue ('Nom');
       DonnerInfoDestinataire (CoordonneeDestinataire,TypeDestinataire,LibelleDestinataire);
       if (LibelleDestinataire='') then TobDest.Detail [Indice].GetValue ('Libelle');

       //--- On ajoute le destinataire si il n'existe pas déjà
       if (UneTob.FindFirst (['NOM'],[CoordonneeDestinataire],True)=nil) then
        begin
         UneTobEnreg:=tob.create ('',UneTob,-1);
         UneTobEnreg.LoadFromSt ('Nom|Libelle|Type','|',CoordonneeDestinataire+'|'+LibelleDestinataire+'|'+TypeDestinataire,'|');
        end;
      end;
    end;

   if (UneTob<>nil) then UneTob.PutGridDetail(GrilleSelection,False,False,'TYPE;NOM');
  end;

 //--- Définition de la présentation de la grille
 Largeur:=GrilleSelection.Width div 5;
 GrilleSelection.ColWidths[0]:=Largeur;      // Type
 GrilleSelection.ColWidths[1]:=4*Largeur;    // Nom
 HMTrad.ResizeGridColumns(GrilleSelection);
end;

//------------------------------------------------------------------
//--- Nom   : InitialiserDestGrilleSelection
//--- Objet : Copie dans la grille de selection les destinataires
//---         choisis dans la grille donnee
//------------------------------------------------------------------
procedure TFRechercherDestinataire.CopierDestGrilleSelection (GrilleSelection,GrilleDonnee : THGrid; Var UneTob : Tob);
var TypeDestinataire       : String;
    CoordonneeDestinataire : String;
    LibelleDestinataire    : String;
    Indice                 : Integer;
    UneTobEnreg            : Tob;
begin
 if (UneTob<>nil) and (GrilleDonnee.Objects [0,1]<>Nil) then
  begin
   TypeDestinataire:=copy (GrilleDonnee.Name,7,3);
   for Indice:=0 to GrilleDonnee.NbSelected-1 do
    begin
     GrilleDonnee.GotoLeBookmark(Indice);
     UneTobEnreg:=Tob(GrilleDonnee.Objects[0,GrilleDonnee.Row]);

     if (UpperCase (TypeDestinataire)='COL') then
      begin
       CoordonneeDestinataire:=UneTobEnreg.GetValue('US_UTILISATEUR');
       LibelleDestinataire:=UneTobEnreg.GetValue('US_LIBELLE');
      end
     else
      if (UpperCase (TypeDestinataire)='ANN') then
       begin
        CoordonneeDestinataire:=UneTobEnreg.GetValue('C_RVA');
        LibelleDestinataire:=UneTobEnreg.GetValue('ANN_NOM');
       end
      else
       if (UpperCase (TypeDestinataire)='REP') then
        begin
         CoordonneeDestinataire:=UneTobEnreg.GetValue('YRP_EMAIL');
         LibelleDestinataire:=Trim (UneTobEnreg.GetValue('YRP_NOM')+' '+UneTobEnreg.GetValue('YRP_PRENOM'));
        end
       else
        if (UpperCase (TypeDestinataire)='LIS') then
         begin
          CoordonneeDestinataire:=UneTobEnreg.GetValue('YLI_LISTEDISTRIB');
          LibelleDestinataire:=UneTobEnreg.GetValue('YLI_LIBELLELISTE');
         end;

     //--- On ajoute le destinataire si il n'existe pas déjà
     if (UneTob.FindFirst (['NOM'],[CoordonneeDestinataire],True)=nil) then
      begin
       UneTobEnreg:=tob.create ('',UneTob,-1);
       UneTobEnreg.LoadFromSt ('Nom|Libelle|Type','|',CoordonneeDestinataire+'|'+LibelleDestinataire+'|'+TypeDestinataire,'|');
      end;
    end;
   GrilleSelection.VidePile (False);
   if (UneTob<>nil) then UneTob.PutGridDetail(GrilleSelection,False,False,'TYPE;NOM');
  end;
end;

//-----------------------------------------------------------------------------
//--- Nom   : SupprimerDestGrilleSelection
//--- Objet : Supprime de la grille selection les destinataires selectionnés
//-----------------------------------------------------------------------------
procedure TFRechercherDestinataire.SupprimerDestGrilleSelection (GrilleSelection : THGrid; Var UneTob : Tob);
var Indice : Integer;
begin
 if (UneTob<>nil) and (GrilleSelection.Objects [0,1]<>Nil) then
  begin
   if GrilleSelection.AllSelected then
    UneTob.ClearDetail
   else
    begin
     for Indice:=GrilleSelection.RowCount-1 Downto 1 do
      begin
       if (GrilleSelection.IsSelected (Indice)) then
        UneTob.detail[Indice-1].Free;
      end;
    end;

   GrilleSelection.VidePile (False);
   GrilleSelection.AllSelected:=False;
   if (UneTob<>nil) then UneTob.PutGridDetail(GrilleSelection,False,False,'TYPE;NOM');
  end;
end;

//---------------------------------------
//--- Nom   : BDestAClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BDestAClick(Sender: TObject);
begin
 inherited;
 if (PageRecherche.ActivePage=OngletCollaborateur) then
  begin
   CopierDestGrilleSelection (GrilleA,GrilleCollaborateur,TobA);
   GrilleCollaborateur.AllSelected:=False;
  end
 else
  if (PageRecherche.ActivePage=OngletAnnuaire) then
   begin
    CopierDestGrilleSelection (GrilleA,GrilleAnnuaire,TobA);
    GrilleAnnuaire.AllSelected:=False;
   end
  else
   if (PageRecherche.ActivePage=OngletRepertoire) then
    begin
     CopierDestGrilleSelection (GrilleA,GrilleRepertoire,TobA);
     GrilleRepertoire.AllSelected:=False;
    end
   else
    if (PageRecherche.ActivePage=OngletListeDistribution) then
     begin
      CopierDestGrilleSelection (GrilleA,GrilleListeDistribution,TobA);
      GrilleListeDistribution.AllSelected:=False;
     end;
end;

//---------------------------------------
//--- Nom   : BDestCcClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BDestCcClick(Sender: TObject);
begin
 inherited;
 if (PageRecherche.ActivePage=OngletCollaborateur) then
  begin
   CopierDestGrilleSelection (GrilleCc,GrilleCollaborateur,TobCc);
   GrilleCollaborateur.AllSelected:=False;
  end
 else
  if (PageRecherche.ActivePage=OngletAnnuaire) then
   begin
    CopierDestGrilleSelection (GrilleCc,GrilleAnnuaire,TobCc);
    GrilleAnnuaire.AllSelected:=False;
   end
  else
   if (PageRecherche.ActivePage=OngletRepertoire) then
    begin
     CopierDestGrilleSelection (GrilleCc,GrilleRepertoire,TobCc);
     GrilleRepertoire.AllSelected:=False;
    end
   else
    if (PageRecherche.ActivePage=OngletListeDistribution) then
     begin
      CopierDestGrilleSelection (GrilleCc,GrilleListeDistribution,TobCc);
      GrilleListeDistribution.AllSelected:=False;
     end;
end;

//-------------------------------------
//--- Nom   : DonnerTypeDestinataire
//--- Objet :
//-------------------------------------
procedure TFRechercherDestinataire.DonnerInfoDestinataire (CoordonneeDestinataire : string; var TypeDestinataire : String; var LibelleDestinataire : String);
Var UneTobEnreg         : Tob;
begin
 TypeDestinataire:='';LibelleDestinataire:='';

 UneTobEnreg:=TobCollaborateur.FindFirst (['US_UTILISATEUR'],[CoordonneeDestinataire],True);
 if (UneTobEnreg<>nil) then
  begin
   TypeDestinataire:='Col';
   LibelleDestinataire:=UneTobEnreg.GetValue ('US_LIBELLE');
   exit;
  end;

 UneTobEnreg:=TobAnnuaire.FindFirst (['C_RVA'],[CoordonneeDestinataire],True);
 if (UneTobEnreg<>nil) then
  begin
   TypeDestinataire:='Ann';
   LibelleDestinataire:=UneTobEnreg.GetValue ('ANN_NOM');
   exit;
  end;

 UneTobEnreg:=TobRepertoire.FindFirst (['YRP_EMAIL'],[CoordonneeDestinataire],True);
 if (UneTobEnreg<>nil) then
  begin
   TypeDestinataire:='Rep';
   LibelleDestinataire:=UneTobEnreg.GetValue ('YRP_NOM');
   exit;
  end;

 if not (ModeAffichageLite) then
  begin
   UneTobEnreg:=TobListeDistribution.FindFirst (['YLI_LISTEDISTRIB'],[CoordonneeDestinataire],True);
   if (UneTobEnreg<>nil) then
    begin
     TypeDestinataire:='Lis';
     LibelleDestinataire:=UneTobEnreg.GetValue ('YLI_LIBELLELISTE');
     exit;
    end;
  end;
end;

//---------------------------------------
//--- Nom   : BValiderClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BValiderClick(Sender: TObject);
begin
 inherited;

 if (TobA<>nil) and (TobDestA<>nil) then
  TobDestA.Dupliquer (TobA,True,True);

 if (TobCc<>nil) and (TobDestCc<>nil) then
  TobDestCc.Dupliquer (TobCc,True,True);
end;

//---------------------------------------
//--- Nom   : BRechercherClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BRechercherClick(Sender: TObject);
begin
 FindDialog.Execute ;
end;

//---------------------------------------
//--- Nom   : BSupprimerDestAClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BSupprimerDestAClick(Sender: TObject);
begin
 inherited;
 SupprimerDestGrilleSelection (GrilleA,TobA);
end;

//---------------------------------------
//--- Nom   : BSupprimerDestCcClick
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.BSupprimerDestCcClick(Sender: TObject);
begin
 inherited;
 SupprimerDestGrilleSelection (GrilleCc,TobCc);
end;

//---------------------------------------
//--- Nom   : FindDialogFind
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.FindDialogFind(Sender: TObject);
var FirstFind : boolean;
begin
 FirstFind:=True;
 case PageRecherche.ActivePageIndex of
  0 : Rechercher(GrilleCollaborateur,FindDialog,FirstFind);
  1 : Rechercher(GrilleAnnuaire,FindDialog,FirstFind);
  2 : Rechercher(GrilleRepertoire,FindDialog,FirstFind);
  3 : Rechercher(GrilleListeDistribution,FindDialog,FirstFind);
 end;
end;

//------------------------------------------------------------
//--- Nom   : RechercherPremiereLettrePremiereColonneGrille
//--- Objet :
//------------------------------------------------------------
procedure TFRechercherDestinataire.RechercherPremiereLettrePremierecolonneGrille (UneGrille : ThGrid;Lettre: String);
var Indice           : Integer;
    IndiceLigneDebut : Integer;
    IndiceRecherche  : Integer;
begin
 //--- recherche à partir de la ligne en cours
 IndiceLigneDebut:=UneGrille.Row+1;

 for Indice:=IndiceLigneDebut to UneGrille.RowCount+IndiceLigneDebut-1 do
  begin
   IndiceRecherche:=Indice;
   //--- On peut partir du milieu de la grille, mais dans ce cas revenir au début
   if IndiceRecherche>UneGrille.RowCount-1 then IndiceRecherche:=IndiceRecherche-UneGrille.RowCount;
   //--- On evite la ligne 0 (les titres)
   if IndiceRecherche=0 then IndiceRecherche:=1;
   //--- On recherche la 1ère lettre de la 1ère colonne
   if UpperCase(Copy(UneGrille.Cells[0,IndiceRecherche],1,1))=Lettre then
    begin
     UneGrille.Row :=IndiceRecherche;
     break;
    end;
  end;
end;

//---------------------------------------
//--- Nom   : FormKeyDown
//--- Objet :
//---------------------------------------
procedure TFRechercherDestinataire.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 inherited;
 case Key of
  //--- Touche Suppr
  VK_DELETE : begin
               if (GrilleA.focused) then SupprimerDestGrilleSelection (GrilleA,TobA);
               if (GrilleCc.focused) then SupprimerDestGrilleSelection (GrilleCc,TobCc);
              end;

  //--- Touche alpha numérique
  48..90 : begin
            if (Shift = [ssCtrl]) then
             begin
              //--- Touche Ctrl + A
              if (Key=65) then
               begin
                if (GrilleCollaborateur.focused) then GrilleCollaborateur.AllSelected:=not (GrilleCollaborateur.AllSelected);
                if (GrilleAnnuaire.focused) then GrilleAnnuaire.AllSelected:=not (GrilleAnnuaire.AllSelected);
                if (GrilleRepertoire.focused) then GrilleRepertoire.AllSelected:=not (GrilleRepertoire.AllSelected);
                if (GrilleListeDistribution.focused) then GrilleListeDistribution.AllSelected:=not (GrilleListeDistribution.AllSelected);
                if (GrilleA.focused) then GrilleA.AllSelected:=not (GrilleA.AllSelected);
                if (GrilleCc.focused) then GrilleCc.AllSelected:=not (GrilleCc.AllSelected);
               end;
              //--- Touche Ctrl + F
              if (Key=70) then
               begin
                BRechercher.Click
               end;
             end;

            //--- Recherche de la 1ere lettre ou chiffre tapé(e)
            if (Shift = []) then
             begin
              if (GrilleCollaborateur.focused) then RechercherPremiereLettrePremierecolonneGrille (GrilleCollaborateur,Chr(Key));
              if (GrilleAnnuaire.focused) then RechercherPremiereLettrePremierecolonneGrille (GrilleAnnuaire,Chr(Key));
              if (GrilleRepertoire.focused) then RechercherPremiereLettrePremierecolonneGrille (GrilleRepertoire,Chr(Key));
              if (GrilleListeDistribution.focused) then RechercherPremiereLettrePremierecolonneGrille (GrilleListeDistribution,Chr(Key));
             end;
           end;
 end;
end;

end.
