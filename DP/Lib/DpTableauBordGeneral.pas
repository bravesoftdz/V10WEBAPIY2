//-------------------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  : Analyse d'activitée : Comptable / Social / GI
//-------------------------------------------------------------
unit DpTableauBordGeneral;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Hctrls, Grids, ComCtrls, uiutil, HPanel,
{$IFDEF EAGLCLIENT}
  UTob, MaineAgl, UtileAgl,
{$ELSE}
  EdtREtat,Utob,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} Fe_Main,
{$ENDIF}
  Hmsgbox, CalcOLEGenericAff, Hent1, ED_Tools, ExtCtrls, Menus, HXlsPas;

//--------------------------------------------------
//--- Déclaration des constantes
//--------------------------------------------------
const
 COMPTA  =  1;
 PAIE    =  2;
 GI      =  3;

type

  TInfoComplementaire = class (TObject)
   CodeAffaire : String;
   Affaire1    : String;
   Affaire2    : String;
   Affaire3    : String;
   Avenant     : String;
   Tiers       : String;
   Ddate       : TDateTime;
   FDate       : TDateTime;
  end;

  TFDPTableauBordGeneral = class(TFVierge)
    Panel2: TPanel;
    GrilleCompta: THGrid;
    Panel3: TPanel;
    HLabel2: THLabel;
    ListeAnnee: THValComboBox;
    GrillePaie: THGrid;
    Panel4: TPanel;
    GrilleGI: THGrid;
    Panel1: TPanel;
    Solde: THLabel;
    LSolde: THLabel;
    MenuOption: TPopupMenu;
    Dtailralise1: TMenuItem;
    Dtailfactur1: TMenuItem;
    CumulAssistant: TMenuItem;
    BEffacerPaie: TToolbarButton97;
    N1: TMenuItem;
    mnGrandLivre: TMenuItem;
    RB_MISSIONSTOUTES: THRadiobutton;
    RB_MISSIONSENCOURS: THRadiobutton;

    procedure FormCreate(Sender: TObject);
    procedure ListeAnneeChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure LSoldeClick(Sender: TObject);
    procedure LSoldeMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure Panel1MouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure FormResize(Sender: TObject);
    procedure CumulAssistantClick(Sender: TObject);
    procedure DetailRealiseClick(Sender: TObject);
    procedure DetailFactureClick(Sender: TObject);
    procedure PurgerComptaClick(Sender: TObject);
    procedure PurgerPaieClick(Sender: TObject);
    procedure PurgerGIClick(Sender: TObject);
    procedure mnGrandLivreClick(Sender: TObject);
    procedure RB_MISSIONSTOUTESClick(Sender: TObject);
    procedure RB_MISSIONSENCOURSClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);

//    procedure GrilleComptaMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
//    procedure GrillePaieMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
//    procedure GrilleGIMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  private
   m_bVoirPR  :boolean;

  public
   NMoins_Existe : Boolean;
   NombreColonnePaie : Integer;
   NumDossier : String;

   TobImprimer : Tob;

   procedure InitialiserTableGI;
   procedure InitialiserSoldeClient (Sender : TObject);
   procedure InitialiserLibelleGrilleCompta (Sender : TObject);
   procedure InitialiserCelluleGrilleCompta (Sender : TObject);
   procedure InitialiserComboBoxPaie (Sender: TObject);
   procedure InitialiserLibelleGrillePaie (Sender : TObject);
   procedure InitialiserCelluleGrillePaie (Sender : TObject);
   procedure InitialiserLibelleGrilleGI (Sender : TObject);
   procedure InitialiserCelluleGrilleGI (Sender : TObject);
   procedure PostDrawCellCompta(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
   procedure PostDrawCellPaie(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
   procedure PostDrawCellGI(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
   function  IsDossierAgricole : Boolean;

   function DonnerInformationCritere : Tob;
   function DonnerInformationGrille (UneGrille : THGrid) : Tob;
   procedure ConstruireTobImprimer (UnLibelleLigne : String; NbreChp : Integer; UneTob : Tob);

  end;

//--------------------------------
//--- Déclaration des fonctions
//--------------------------------
function CalculerSolde (NumDossier : String; AfficherErreur : Boolean=True) : Double;

//---------------------------------
//--- Déclaration des procédures
//---------------------------------
procedure Aff_TableauBordGeneral (Dossier : String);
procedure ConsulterSolde (NumDossier : String);

implementation

uses {DpTableauBordLibrairie, AfTableauBord, UtofAftableauBord, uTofConsEcr,
     DpTableauBordDetailSolde, DpTableauBordAssistant, AnnOutils, Ent1, AffaireUtil,}

{$IFDEF BUREAU} // 3OB 27/08/07
DpTableauBordDetailSolde,
DpTableauBordAssistant,
uTofConsEcr,
{$ENDIF}

// 3OB 27/08/07  : fin

DpTableauBordLibrairie,
AfTableauBord, UtofAftableauBord,
  AnnOutils, Ent1, AffaireUtil,
     ConfidentAffaire, // $$$JP 04/08/05 - pour voir PR ou non
     GalOutil, GalSystem;

{$R *.DFM}

//---------------------------------------------------------
//--- Nom   : Aff_tableauBordGeneral
//--- Objet : Affichage de la fiche Tableau bord général
//---------------------------------------------------------
procedure Aff_TableauBordGeneral (Dossier : String);
Var Fiche : TFDpTableauBordGeneral;
    PP : THPanel;
    NumLigne : Integer;
begin
 Fiche:=TFDpTableauBordGeneral.Create (Application);
 Fiche.NumDossier:=Dossier;
 PP:=FindInsidePanel;
 if PP=nil then
  begin
   try
    Fiche.showModal;
   finally;
    //--- Libération des objects grille GI
    for NumLigne:=1 to Fiche.GrilleGI.RowCount do
     Fiche.GrilleGI.Objects [1,NumLigne].Free;

    Fiche.Free;
   end;
  end
 else
  begin
   InitInside (Fiche,PP);
   Fiche.Show;
  end;
end;

//-----------------------------------
//--- Nom   : FormCreate
//--- Objet : Création de la fiche
//-----------------------------------
procedure TFDPTableauBordGeneral.FormCreate(Sender: TObject);
begin
 inherited;
 GrilleCompta.Ctl3D:=False;
 GrillePaie.Ctl3D:=False;
 GrilleGI.Ctl3D:=False;
 GrilleCompta.PostDrawCell:=PostDrawCellCompta;
 GrillePaie.PostDrawCell:=PostDrawCellPaie;
 GrilleGI.PostDrawCell:=PostDrawCellGI;
 Self.Caption:='Analyse activité';

 // $$$ JP 04/08/05 - ne pas voir les PR si pas le droit
 m_bVoirPR := AffichageValorisation;
 if m_bVoirPR = FALSE then
 begin
      Dtailralise1.Enabled := FALSE;
      mnGrandLivre.Enabled := FALSE;
 end;
end;

//------------------------------------
//--- Nom   : FormShow
//--- Objet : Affichage de la fiche
//------------------------------------
procedure TFDPTableauBordGeneral.FormShow(Sender: TObject);
begin
 inherited;
 InitMoveProgressForm(nil, TitreHalley,'Intégration des données',3, False, True);
 TextProgressForm ('Travail en cours ... Veuillez patienter ...');

// 3OB 27/08/07 : cacher la compta et la paie si on vient de la GI
{$IFDEF BUREAU}
 MoveCurProgressForm('Comptable');
 Grillecompta.EnabledBlueButton:=True;
 InitialiserLibelleGrilleCompta (Sender);
 InitialiserCelluleGrilleCompta (Sender);

 MoveCurProgressForm('Paie');
 GrillePaie.EnabledBlueButton:=True;
 InitialiserComboBoxPaie (Sender);
 InitialiserLibelleGrillePaie (Sender);
 InitialiserCelluleGrillePaie (Sender);

   Grillecompta.visible:=true;
   GrillePaie.visible:=true;
   Panel1.visible:=true;
   Panel2.visible:=true;
   Panel3.visible:=true;
{$ELSE}
   Grillecompta.visible:=false;
   GrillePaie.visible:=false;
   Panel1.visible:=false;
   Panel2.visible:=false;
   Panel3.visible:=false;
   BeffacerPaie.visible:=false;
   mnGrandLivre.visible := FALSE;
// 3OB 27/08/07  : Fin
{$ENDIF}
 MoveCurProgressForm('Gestion Interne');

 if (IsCodeTiersExiste (NumDossier)) and (not IsDossierCabinet (NumDossier)) then
  begin
   InitialiserTableGI;
   GrilleGI.EnabledBlueButton:=True;
   InitialiserLibelleGrilleGI (Sender);
   InitialiserCelluleGrilleGI (Sender);
  end
 else
  begin
   GrilleGI.EnabledBlueButton:=True;
   InitialiserLibelleGrilleGI (Sender);
  end;
{$IFDEF BUREAU}  // 3OB 27/08/07 
 if not ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_NODOSSIER="'+NumDossier+'" AND DOS_CABINET="X"') then
  InitialiserSoldeClient (Sender)
 else
  begin
   Solde.Visible := False;
   LSolde.Visible := False;
  end;
{$ENDIF}  // 3OB 27/08/07
 FiniMoveProgressForm;
end;

//--------------------------------------------------------------
//--- Nom   : InitialiserLibelleGrilleCompta
//--- Objet : Initialisation des libelles de la grille compta
//--------------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserLibelleGrilleCompta (Sender : TObject);
var SSql            : String;
    QDos            : Tquery;
    NumColonne      : Integer;
    Indice          : Integer;
begin
 GrilleCompta.ColCount:=4;
 GrilleCompta.RowCount:=5;
 GrilleCompta.Height:=100;
 GrilleCompta.ColWidths [0]:=100;
 GrilleCompta.Cells [0,1]:='Nombre écriture(s)';
 GrilleCompta.Cells [0,2]:='Nombre immo(s)';
 GrilleCompta.Cells [0,3]:='Nombre entrée(s)';
 GrilleCompta.Cells [0,4]:='Nombre sortie(s)';

 if (IsDossierAgricole) then
  begin
   GrilleCompta.RowCount:=9;
   GrilleCompta.Height:=180;
   GrilleCompta.Cells [0,5]:='Subv produits';
   GrilleCompta.Cells [0,6]:='Subv structures';
   GrilleCompta.Cells [0,7]:='Subv revenus';
   GrilleCompta.Cells [0,8]:='Autres subv';
  end;

 SSql:='SELECT * from DPTABCOMPTA where DTC_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE, -1, '', True);
   for Indice:=1 to GrilleCompta.ColCount-1 do
    begin
     GrilleCompta.ColWidths [Indice]:=((GrilleCompta.Width-100) div 3);
     if (not QDos.Eof) then
      begin
       //--- Determine le numéro de la colonne
       if (QDos.FindField ('DTC_MILLESIME').AsString='N-') then
        NumColonne:=1
       else
        if (QDos.FindField ('DTC_MILLESIME').AsString='N') then
         NumColonne:=2
        else
         NumColonne:=3;

       //--- Affichage des libelles colonnes
       GrilleCompta.Cells [NumColonne,0]:=QDos.FindField ('DTC_LIBEXERCICE').AsString;

       QDos.Next;
      end;
    end;
   Ferme (QDos);
  except
   begin
    PGIINFO ('Problème dans la table DPTABCOMPTA','Erreur');
   end;
  end;
end;

//--------------------------------------------------------------
//--- Nom   : InitialiserCelluleGrilleCompta
//--- Objet : Initialisation des cellules de la grille compta
//--------------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserCelluleGrilleCompta (Sender : TObject);
var SSql       : String;
    QDos       : Tquery;
    NumColonne : Integer;
    Indice     : Integer;
begin
 NMoins_Existe:=False;
 SSql:='SELECT * from DPTABCOMPTA where DTC_NODOSSIER="'+NumDossier+'"';
  try
   QDos:=Opensql (SSql, TRUE, -1, '', True);
   for Indice:=1 to GrilleCompta.ColCount-1 do
    begin
     if (not QDos.Eof) then
      begin
       //--- Détermine le numéro de la colonne
       if (QDos.FindField ('DTC_MILLESIME').AsString='N-') then
        begin
         NMoins_Existe:=True;
         NumColonne:=1
        end
       else
        if (QDos.FindField ('DTC_MILLESIME').AsString='N') then
         NumColonne:=2
        else
         NumColonne:=3;

       GrilleCompta.Cells [NumColonne,1]:=FormaterNombre (QDos.FindField ('DTC_NBECRITURE').AsString);
       GrilleCompta.Cells [NumColonne,2]:=FormaterNombre (QDos.FindField ('DTC_NBLIGNEIMMO').AsString);
       GrilleCompta.Cells [NumColonne,3]:=FormaterNombre (QDos.FindField ('DTC_NBENTREEIMMO').AsString);
       GrilleCompta.Cells [NumColonne,4]:=FormaterNombre (QDos.FindField ('DTC_NBSORTIEIMMO').AsString);

       if (IsDossierAgricole) then
        begin
         GrilleCompta.Cells [NumColonne,5]:=StrFMontant (QDos.FindField ('DTC_SUBVPRO').AsFloat,15,2,'',True);
         GrilleCompta.Cells [NumColonne,6]:=StrFMontant (QDos.FindField ('DTC_SUBVSTR').AsFloat,15,2,'',True);
         GrilleCompta.Cells [NumColonne,7]:=StrFMontant (QDos.FindField ('DTC_SUBVREV').AsFloat,15,2,'',True);
         GrilleCompta.Cells [NumColonne,8]:=StrFMontant (QDos.FindField ('DTC_SUBVAUT').AsFloat,15,2,'',True);
        end;

       QDos.Next;
      end;
    end;

   Ferme (QDos);
  except
   begin
    PGIINFO ('Problème dans la table DPTABCOMPTA','Erreur');
   end;
  end;
end;

//----------------------------------------------------------
//--- Nom   : InitialiserComboBoxPaie
//--- Objet : Initialisation des éléments de la combo box
//----------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserComboBoxPaie (Sender: TObject);
var SSql       : String;
    QDos       : Tquery;
begin
 SSql:='Select DT1_ANNEE from DPTABGENPAIE where DT1_NODOSSIER="'+NumDossier+'" group by DT1_ANNEE';
  try
   QDos:=Opensql (SSql, TRUE, -1, '', True);

   While Not QDos.Eof do
    begin
     ListeAnnee.Items.AddObject (QDos.FindField ('DT1_ANNEE').AsString,nil);
     QDos.Next;
    end;
   if (ListeAnnee.Items.Count>0) then ListeAnnee.ItemIndex:=ListeAnnee.Items.Count-1;
   Ferme (QDos);
  except
   begin
    PGIINFO ('Problème dans la table DPTABGENPAIE','Erreur');
   end;
  end;
end;

//--------------------------------------------------------------
//--- Nom   : InitialiserLibelleGrillePaie
//--- Objet : Initialisation des libelles de la grille Paie
//--------------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserLibelleGrillePaie (Sender : TObject);
var Indice : Integer;
begin
 GrillePaie.ColCount:=14;
 GrillePaie.RowCount:=6;
 GrillePaie.ColWidths [0]:=100;

 GrillePaie.Cells [0,1]:='Nombre Entrée(s)';
 GrillePaie.Cells [0,2]:='Nombre Sortie(s)';
 GrillePaie.Cells [0,3]:='Nombre Salarié(s)';
 GrillePaie.Cells [0,4]:='Nombre Bulletin(s)';
 GrillePaie.Cells [0,5]:='Total Brut';

 for Indice:=1 to 13 do
  GrillePaie.ColWidths [Indice]:=((GrillePaie.Width-100) div 7)-1;

 GrillePaie.Cells [1,0]:='Total';
 GrillePaie.Cells [2,0]:='Janvier';
 GrillePaie.Cells [3,0]:='Février';
 GrillePaie.Cells [4,0]:='Mars';
 GrillePaie.Cells [5,0]:='Avril';
 GrillePaie.Cells [6,0]:='Mai';
 GrillePaie.Cells [7,0]:='Juin';
 GrillePaie.Cells [8,0]:='Juillet';
 GrillePaie.Cells [9,0]:='Août';
 GrillePaie.Cells [10,0]:='Septembre';
 GrillePaie.Cells [11,0]:='Octobre';
 GrillePaie.Cells [12,0]:='Novembre';
 GrillePaie.Cells [13,0]:='Décembre';
end;

//------------------------------------------------------------
//--- Nom   : InitialiserCelluleGrillePaie
//--- Objet : Initialisation des cellules de la grille Paie
//------------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserCelluleGrillePaie (Sender : TObject);
var SSql                        : String;
    QDos                        : Tquery;
    Indice                      : Integer;
    TotalEntree, TotalSortie    : Integer;
    TotalPresent, TotalBulletin : Integer;
    TotalBrut                   : Double;
begin
 NMoins_Existe:=False;
 SSql:='SELECT * FROM DPTABGENPAIE WHERE DT1_NODOSSIER="'+NumDossier+'" and DT1_ANNEE="'+ListeAnnee.Text+'" ORDER BY DT1_MOIS';
  try
   QDos:=Opensql (SSql, TRUE, -1, '', True);

   TotalEntree:=0;TotalSortie:=0;TotalPresent:=0;TotalBulletin:=0;TotalBrut:=0;
   for Indice:=1 to GrillePaie.ColCount-1 do
    begin
     if (not QDos.Eof) then
      begin
       NombreColonnePaie:=StrToInt (QDos.FindField ('DT1_MOIS').AsString);

       GrillePaie.Cells [NombreColonnePaie+1,1]:=FormaterNombre (QDos.FindField ('DT1_NBENTREES').AsString);
       TotalEntree:=TotalEntree+StrToInt (QDos.FindField ('DT1_NBENTREES').AsString);
       GrillePaie.Cells [NombreColonnePaie+1,2]:=FormaterNombre (QDos.FindField ('DT1_NBSORTIES').AsString);
       TotalSortie:=TotalSortie+StrToInt (QDos.FindField ('DT1_NBSORTIES').AsString);
       GrillePaie.Cells [NombreColonnePaie+1,3]:=FormaterNombre (QDos.FindField ('DT1_NBPRESENTS').AsString);
       TotalPresent:=TotalPresent+StrToInt (QDos.FindField ('DT1_NBPRESENTS').AsString);
       GrillePaie.Cells [NombreColonnePaie+1,4]:=FormaterNombre (QDos.FindField ('DT1_NBBULLETINS').AsString);
       TotalBulletin:=TotalBulletin+StrToInt (QDos.FindField ('DT1_NBBULLETINS').AsString);
       GrillePaie.Cells [NombreColonnePaie+1,5]:=StrFMontant (QDos.FindField ('DT1_TOTALBRUT').AsFloat,15,2,'',True);
       TotalBrut:=TotalBrut+QDos.FindField ('DT1_TOTALBRUT').AsFloat;

       QDos.Next;
      end;
    end;

   //--- Affiche la colonne Total
   GrillePaie.Cells [1,1]:=FormaterNombre (IntToStr (TotalEntree));
   GrillePaie.Cells [1,2]:=FormaterNombre (IntToStr (TotalSortie));
   GrillePaie.Cells [1,3]:=FormaterNombre (IntToStr (TotalPresent));
   GrillePaie.Cells [1,4]:=FormaterNombre (IntToStr (TotalBulletin));
   GrillePaie.Cells [1,5]:=StrFMontant (TotalBrut,15,2,'',True);

   Ferme (QDos);

  except
   begin
    PGIINFO ('Problème dans la table DPTABGENPAIE','Erreur');
   end;
  end;
end;

//----------------------------------------------------------
//--- Nom   : InitialiserTableGI
//--- Objet : Initialisation de la table GI AFTableauBord
//----------------------------------------------------------
procedure TFDPTableauBordGeneral.InitialiserTableGI;
Var Param : T_PARAMTB;
begin
 InitRecordParamTb(Param);
 Param.AlimPRevu:=true;
 Param.Cltdeb:=DonnerTiersFromGuid (DonnerGuidFromDossier (NumDossier));
 Param.CltFin:=DonnerTiersFromGuid (DonnerGuidFromDossier (NumDossier));
 Param.DateDeb :=StrToDate('01/01/1900');
 Param.DateFin :=StrToDate('31/12/2099');
 if not AfBuildAfTableauBord(Param) then
  PGIINFO ('Problème dans la table AFTableauBord','Erreur');
end;

//----------------------------------------------------------
//--- Nom   : InitialiserLibelleGrilleGI
//--- Objet : Initialisation des libelles de la grille GI
//----------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserLibelleGrilleGI (Sender : TObject);
var
   Indice         :integer;
   LargeurColonne :integer;
begin
 GrilleGI.RowCount:=2;
 GrilleGI.ColCount:=12;
 GrilleGI.ColWidths [0]:=100;
 GrilleGI.ColWidths [1]:=130;
 LargeurColonne:=(GrilleGI.Width - (GrilleGI.ColWidths [0]+GrilleGI.ColWidths [1])) div (GrilleGI.ColCount-2);

 for Indice:=2 to GrilleGI.ColCount-1 do
  GrilleGI.ColWidths [Indice]:=LargeurColonne;

 //GrilleGI.Cells [0,0]:='Mission';
 GrilleGI.Cells [1,0]:='Responsable';
 GrilleGI.Cells [2,0]:='Prévu (h)';
 GrilleGI.Cells [3,0]:='Réalisé (h)';
 GrilleGI.Cells [4,0]:='Rea/Prév (h)';
 GrilleGI.Cells [5,0]:='Prévu';
 GrilleGI.Cells [6,0]:='Réalisé Pr';
 GrilleGI.Cells [7,0]:='Réalisé Pv';
 GrilleGI.Cells [8,0]:='Facturé';
 GrilleGI.Cells [9,0]:='Fac/Pv';
 GrilleGI.Cells [10,0]:='Fac/Pr';
 GrilleGI.Cells [11,0]:='Fac/Prév';

 HMTrad.ResizeGridColumns(GrilleGI);
end;

//----------------------------------------------------------
//--- Nom   : InitialiserCelluleGrilleGI
//--- Objet : Initialisation des cellules de la grille GI
//----------------------------------------------------------
procedure TFDpTableauBordGeneral.InitialiserCelluleGrilleGI (Sender : TObject);
var
 SSql,Mission,Responsable                                        : String;
 QDos,RSql                                                       : Tquery;
 NumLigne                                                        : Integer;
 BudgetHeure,HeureRealise,EcartBudRea                            : Double;
 BudgetMontant,RealisePr,RealisePv,Facture                       : Double;
 EcartFacPv,EcartFacPr,EcartFacBud                               : Double;
 TotalBudgetHeure,TotalHeureRealise,TotalEcartBudRea             : Double;
 TotalBudgetMontant,TotalRealisePr,TotalRealisePv,TotalFacture   : Double;
 TotalEcartFacPv,TotalEcartFacPr,TotalEcartFacBud                : Double;
 CodeTier                                                        : String;
 InfoComplementaire                                              : TInfoComplementaire;
begin
 NumLigne:=1;
 BudgetHeure:=0;HeureRealise:=0;EcartBudRea:=0;
 BudgetMontant:=0;RealisePr:=0;RealisePv:=0;Facture:=0;
 EcartFacPv:=0;EcartFacPr:=0;EcartFacBud:=0;
 TotalBudgetHeure:=0;TotalHeureRealise:=0;TotalEcartBudRea:=0;
 TotalBudgetMontant:=0;TotalRealisePr:=0;TotalRealisePv:=0;TotalFacture:=0;
 TotalEcartFacPv:=0;TotalEcartFacPr:=0;TotalEcartFacBud:=0;

 CodeTier:=DonnerTiersFromGuid(DonnerGuidFromDossier (NumDossier));

 if (CodeTier<>'') then
  begin
   SSql:='SELECT ATB_AFFAIRE, ATB_AFFAIRE1, ATB_AFFAIRE2, ATB_AFFAIRE3, ATB_AVENANT,';
   SSql:=SSql+' ATB_TIERS, ATB_DATEDEBUT, ATB_DATEFIN,';
   SSql:=SSql+' max (AFF_ETATAFFAIRE) AS AFF_ETATAFFAIRE, SUM (ATB_QTEUNITEREF) AS ATB_QTEUNITEREF,';
   SSql:=SSql+' SUM (ATB_TOTPRCHARGE) AS ATB_TOTPRCHARGE, SUM (ATB_TOTVENTEACT) AS ATB_TOTVENTEACT, SUM (ATB_FACPV) AS ATB_FACPV,';
   SSql:=SSql+' SUM (ATB_PREVUPR) AS ATB_PREVUPR, SUM (ATB_PREVUPV) AS ATB_PREVUPV, SUM (ATB_PREQTEUNITEREF) AS ATB_PREQTEUNITEREF, ATB_RESPONSABLE';
   if (RB_MISSIONSENCOURS.Checked) then
    SSql:=SSql+' FROM AFFAIRE,AFTABLEAUBORD WHERE AFF_AFFAIRE=ATB_AFFAIRE and AFF_ETATAFFAIRE="ENC" AND AFF_AFFAIRE0="A" AND'
   else
    SSql:=SSql+' FROM AFFAIRE,AFTABLEAUBORD WHERE AFF_AFFAIRE=ATB_AFFAIRE AND AFF_AFFAIRE0="A" AND';
   SSql:=SSql+' ATB_UTILISATEUR="'+V_PGI.USER+'" AND ATB_TIERS="'+CodeTier+'" GROUP BY ATB_AFFAIRE,ATB_RESPONSABLE,';
   SSql:=SSql+' ATB_AFFAIRE1, ATB_AFFAIRE2, ATB_AFFAIRE3, ATB_AVENANT, ATB_TIERS, ATB_DATEDEBUT, ATB_DATEFIN';

   try
    QDos:=Opensql (SSql, TRUE);
    QDos.First;

    if (QDos.RecordCount=0) then
     GrilleGI.RowCount:=2
    else
     GrilleGI.RowCount:=QDos.RecordCount+2;

    while (not QDos.Eof) do
     begin
      //--- Calcul la valeur des colonnes
      Mission:=CodeAffaireAffiche (QDos.FindField ('ATB_AFFAIRE').AsString,' ');
      Responsable:=QDos.FindField ('ATB_RESPONSABLE').AsString+' '+RechDom ('AFLRESSOURCE',QDos.FindField ('ATB_RESPONSABLE').AsString,False);

      BudgetHeure:=QDos.FindField ('ATB_PREQTEUNITEREF').AsFloat;
      HeureRealise:=QDos.FindField ('ATB_QTEUNITEREF').AsFloat;
      EcartBudRea:=HeureRealise-BudgetHeure;
      TotalBudgetHeure:=TotalBudgetHeure+BudgetHeure;
      TotalHeureRealise:=TotalHeureRealise+HeureRealise;
      TotalEcartBudRea:=TotalEcartBudRea+EcartBudRea;

      BudgetMontant:=QDos.FindField ('ATB_PREVUPV').AsFloat;
      RealisePr:=QDos.FindField ('ATB_TOTPRCHARGE').AsFloat;
      RealisePv:=QDos.FindField ('ATB_TOTVENTEACT').AsFloat;
      Facture:=QDos.FindField ('ATB_FACPV').AsFloat;
      TotalBudgetMontant:=TotalBudgetMontant+BudgetMontant;
      TotalRealisePr:=TotalRealisePr+RealisePr;
      TotalRealisePv:=TotalRealisePv+RealisePv;
      TotalFacture:=TotalFacture+Facture;

      EcartFacPv:=Facture-RealisePv;
      EcartFacPr:=Facture-RealisePr;
      EcartFacBud:=Facture-BudgetMontant;
      TotalEcartFacPv:=TotalEcartFacPv+EcartFacPv;
      TotalEcartFacPr:=TotalEcartFacPr+EcartFacPr;
      TotalEcartFacBud:=TotalEcartFacBud+EcartFacBud;

      //--- Affichage des valeurs des colonnes
      if (QDos.FindField ('AFF_ETATAFFAIRE').AsString='ENC') then
       GrilleGI.Cells [0,NumLigne]:=Mission+' [ENC]'
      else
       GrilleGI.Cells [0,NumLigne]:=Mission;
      GrilleGI.Cells [1,NumLigne]:=Responsable;
      GrilleGI.Cells [2,NumLigne]:=StrFMontant (BudgetHeure,15,2,'',True);
      GrilleGI.Cells [3,NumLigne]:=StrFMontant (HeureRealise,15,2,'',True);
      GrilleGI.Cells [4,NumLigne]:=StrFMontant (EcartBudRea,15,2,'',True);
      GrilleGI.Cells [5,NumLigne]:=StrFMontant (BudgetMontant,15,2,'',True);
      if m_bVoirPR = TRUE then
          GrilleGI.Cells [6,NumLigne]:=StrFMontant (RealisePr,15,2,'',True);
      GrilleGI.Cells [7,NumLigne]:=StrFMontant (RealisePv,15,2,'',True);
      GrilleGI.Cells [8,NumLigne]:=StrFMontant (Facture,15,2,'',True);
      GrilleGI.Cells [9,NumLigne]:=StrFMontant (EcartFacPv,15,2,'',True);
      if m_bVoirPR = TRUE then
         GrilleGI.Cells [10,NumLigne]:=StrFMontant (EcartFacPr,15,2,'',True);
      GrilleGI.Cells [11,NumLigne]:=StrFMontant (EcartFacBud,15,2,'',True);

      //--- Ajout des informations complementaires
      InfoComplementaire:=TInfoComplementaire.Create;
      InfoComplementaire.CodeAffaire:=QDos.FindField ('ATB_AFFAIRE').AsString;
      InfoComplementaire.Affaire1:=QDos.FindField ('ATB_AFFAIRE1').AsString;
      InfoComplementaire.Affaire2:=QDos.FindField ('ATB_AFFAIRE2').AsString;
      InfoComplementaire.Affaire3:=QDos.FindField ('ATB_AFFAIRE3').AsString;
      InfoComplementaire.Avenant:=QDos.FindField ('ATB_AVENANT').AsString;
      InfoComplementaire.Tiers:=QDos.FindField ('ATB_TIERS').AsString;
      InfoComplementaire.Ddate:=QDos.FindField ('ATB_DATEDEBUT').AsDateTime;
      InfoComplementaire.Fdate:=QDos.FindField ('ATB_DATEFIN').AsDateTime;

      GrilleGI.Objects [1,NumLigne]:=InfoComplementaire;

      QDos.Next;
      Inc (NumLigne);
     end;

    //--- Affichage de la ligne Total Client
    GrilleGI.Cells [0,NumLigne]:='Total Client';
    GrilleGI.Cells [2,NumLigne]:=StrFMontant (TotalBudgetHeure,15,2,'',True);
    GrilleGI.Cells [3,NumLigne]:=StrFMontant (TotalHeureRealise,15,2,'',True);
    GrilleGI.Cells [4,NumLigne]:=StrFMontant (TotalEcartBudRea,15,2,'',True);
    GrilleGI.Cells [5,NumLigne]:=StrFMontant (TotalBudgetMontant,15,2,'',True);
    if m_bVoirPR = TRUE then
       GrilleGI.Cells [6,NumLigne]:=StrFMontant (TotalRealisePr,15,2,'',True);
    GrilleGI.Cells [7,NumLigne]:=StrFMontant (TotalRealisePv,15,2,'',True);

    if (RB_MISSIONSTOUTES.Checked) then
     begin
      SSql:='SELECT SUM (GL_TOTALHT) AS SOMMEFACTSSDOS FROM LIGNE WHERE GL_AFFAIRE="" AND GL_TIERS="'+CodeTier+'"';
      RSql:=Opensql (SSql, TRUE);
      GrilleGI.Cells [8,NumLigne]:=StrFMontant (TotalFacture+RSql.FindField ('SOMMEFACTSSDOS').AsInteger,15,2,'',True);
      Ferme (RSql)
     end
    else
     GrilleGI.Cells [8,NumLigne]:=StrFMontant (TotalFacture,15,2,'',True);

    GrilleGI.Cells [9,NumLigne]:=StrFMontant (TotalEcartFacPv,15,2,'',True);
    if m_bVoirPR = TRUE then
      GrilleGI.Cells [10,NumLigne]:=StrFMontant (TotalEcartFacPr,15,2,'',True);
    GrilleGI.Cells [11,NumLigne]:=StrFMontant (TotalEcartFacBud,15,2,'',True);

    //--- Ajout des informations complementaires pour la ligne de Totaux
    InfoComplementaire:=TInfoComplementaire.Create;
    InfoComplementaire.CodeAffaire:='###TOTAUX###';
    InfoComplementaire.Tiers:=CodeTier;
    InfoComplementaire.Ddate:=iDate1900;
    InfoComplementaire.Fdate:=iDate2099;

    GrilleGI.Objects [1,NumLigne]:=InfoComplementaire;

    Ferme (QDos);

   except
    begin
     PGIINFO ('Problème dans la table AFTABLEAUBORD','Erreur');
    end;
   end;
  end
 else
  PGIINFO ('La fiche annuaire n''est pas liée à un tiers','Erreur');

end;

//-----------------------------------------
//--- Nom   : InitialiserSoldeClient
//--- Objet : Initialise le solde client
//-----------------------------------------
procedure TFDPTableauBordGeneral.InitialiserSoldeClient (Sender : TObject);
begin
     Solde.Caption:=FormaterMontant (CalculerSolde (NumDossier));
end;

//------------------------------------------------
//--- Nom   : LSoldeClick
//--- Objet : Click sur le label "Solde Client"
//------------------------------------------------
procedure TFDPTableauBordGeneral.LSoldeClick(Sender: TObject);
begin
 inherited;
 LSolde.Color:=ClBtnFace;
 ConsulterSolde (NumDossier);
end;

//------------------------------
//--- Nom   : CalculerSolde
//--- Objet : Calcul du solde
//------------------------------
function CalculerSolde (NumDossier : String; AfficherErreur : Boolean=True) : Double;
var ChSql       : String;
    RSql,RSql2  : TQuery;
    SoldeClient : Double;
    CodeAuxi    : String;
begin
 ChSql:='SELECT DOS_NODOSSIER from DOSSIER where DOS_CABINET="X"';
 SoldeClient:=0;
 CodeAuxi := DonnerAuxiFromGuid(DonnerGuidFromDossier (NumDossier));
 try
  RSql:=Opensql (ChSql, TRUE);
  RSql.First;

  while not (RSql.eof) do
   begin
    if not IsDBZero (RSql.FindField ('DOS_NODOSSIER').AsString) then
     begin
      if not DBExists('DB'+RSql.FindField ('DOS_NODOSSIER').AsString) then
        begin
        PGIInfo('Le solde client ne peut être affiché correctement puisque le dossier '+RSql.FindField ('DOS_NODOSSIER').AsString+' est détaché', TitreHalley);
        SoldeClient := 0;
        break;
        end;
      ChSql:='SELECT T_TOTALCREDIT,T_TOTALDEBIT from DB'+RSql.FindField ('DOS_NODOSSIER').AsString+'.dbo.TIERS Where T_AUXILIAIRE="'+CodeAuxi+'"';
      try
       RSql2:=Opensql (ChSql, TRUE);
       if (Not RSql2.Eof) then
        SoldeClient:=SoldeClient+RSql2.FindField ('T_TOTALDEBIT').AsFloat-RSql2.FindField ('T_TOTALCREDIT').AsFloat;
        // $$$ JP 07/01/2005 - pourquoi réinitialiser le cumul? (bug je pense)
//       else
   //     SoldeClient:=0;
       Ferme (RSql2);
      except
       if (AfficherErreur) then
        PGIINFO ('Problème dans la table TIERS','Erreur');
      end;
     end
    else
     begin
      if (RSql.RecordCount=1) then
       begin
        ChSql:='SELECT T_TOTALCREDIT,T_TOTALDEBIT from TIERS Where T_AUXILIAIRE="'+CodeAuxi+'"';
        try
         RSql2:=Opensql (ChSql, TRUE);
         if (Not RSql2.Eof) then
          SoldeClient:=SoldeClient+RSql2.FindField ('T_TOTALDEBIT').AsFloat-RSql2.FindField ('T_TOTALCREDIT').AsFloat;
        // $$$ JP 07/01/2005 - pourquoi réinitialiser le cumul? (bug, ou plutôt inutile)
//         else
   //       SoldeClient:=0;
         Ferme (RSql2);
        except
         if (AfficherErreur) then
          PGIINFO ('Problème dans la table TIERS','Erreur');
        end;
       end;
     end;
    RSql.Next;
   end;

  Ferme (RSql);
 except
  if (AfficherErreur) then
   PGIINFO ('Problème dans la table DOSSIER','Erreur');
 end;
 CalculerSolde:=SoldeClient;
end;

//------------------------------------
//--- Nom   : ConsulterSolde
//--- Objet : Consultation du solde
//------------------------------------
procedure ConsulterSolde (NumDossier : String);
var ChSql : String;
    RSql  : TQuery;
    CodeAuxiliaire : String;
    CodeGenerale : String;
    ChDossier : String;
begin
 if Not JaiLeDroitConceptBureau(ccAccesGlobalCompta) then exit;

 ChSql:='Select DOS_NODOSSIER from DOSSIER where DOS_CABINET="X"';
 try
  RSql:=Opensql (ChSql, TRUE);
  RSql.First;

  if not (RSql.Eof) then
   begin
    ChDossier := RSql.FindField ('DOS_NODOSSIER').AsString;
    if (RSql.RecordCount=1) and (IsDBZero (ChDossier)) then
     begin
      CodeAuxiliaire:=DonnerAuxiFromGuid (DonnerGuidFromDossier (NumDossier));
      CodeGenerale:=DonnerGenFromAuxi (CodeAuxiliaire);
{$IFDEF BUREAU}   // 3OB 27/08/07 
      uTofConsEcr.OperationsSurComptes (CodeGenerale,'-2','',CodeAuxiliaire,False,'DB'+ChDossier);
{$ENDIF}   // 3OB 27/08/07
     end
    else
{$IFDEF BUREAU}   // 3OB 27/08/07
     AfficherDetailSoldeClient (NumDossier);
{$ENDIF}   // 3OB 27/08/07 
   end;

  Ferme (RSql);
 except
  PGIINFO ('Problème dans la table Dossier','Erreur');
 end;
end;

//----------------------------------------------
//--- Nom   : FormResize
//--- Objet : Redimensionnement de la fenêtre
//----------------------------------------------
procedure TFDPTableauBordGeneral.FormResize(Sender: TObject);
begin
// InitialiserLibelleGrilleCompta (Sender);
// InitialiserLibelleGrillePaie (Sender);
// InitialiserLibelleGrilleGI (Sender);
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu de la cellule avant que la cellule soit
Suite ........ : dessiné par le canvas
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBordGeneral.PostDrawCellCompta(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
 if (ACol>0) and (ARow>0) then
  EcrireTexte (CENTRE,ACol,ARow,GrilleCompta);

 if (not NMoins_Existe) and (ACol=1) then
  GriserGrille (Acol,ARow,GrilleCompta);
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu de la cellule avant que la cellule soit
Suite ........ : dessiné par le canvas
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBordGeneral.PostDrawCellPaie(ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
 if (ACol>0) and (ARow>0) then
   EcrireTexte (DROITE,ACol,ARow,GrillePaie);

 if (GrillePaie.Cells[ACol,ARow]='') then
  GriserGrille (ACol,Arow,GrillePaie);
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 11/06/2003
Modifié le ... :   /  /
Description .. : Définit le contenu de la cellule avant que la cellule soit
Suite ........ : dessiné par le canvas
Suite ........ : $$$ JP 04/08/05 - grisé cellule de type PR si pas le droit de voir
Mots clefs ... :
*****************************************************************}
procedure TFDpTableauBordGeneral.PostDrawCellGI (ACol,ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
   iAlign   :integer;
begin
     // Par défaut, on ne veut pas d'affichage de la cellule
     iAlign := -1;
     if ARow = 0 then
        iAlign := CENTRE
     else
         if aCol = 0 then
              iAlign := GAUCHE
         else if aCol = 1 then
              iAlign := CENTRE
         else if ((aCol <> 6) and (aCol <> 10)) or (m_bVoirPR = TRUE) then
              iAlign := DROITE;

     // Ecriture du texte aligné, ou grisage si demandé
     if iAlign <> -1 then
         EcrireTexte (iAlign, ACol, ARow, GrilleGI)
     else
         GriserGrille (ACol, ARow, GrilleGI);
{ if (ARow>0) and (ACol=1) then
  EcrireTexte (CENTRE,ACol,ARow,GrilleGI);
 if (ARow>0) and (ACol=0) then
  EcrireTexte (Gauche,ACol,ARow,GrilleGI);
 if (ARow>0) and (ACol>1) then
  EcrireTexte (DROITE,ACol,ARow,GrilleGI);}
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 23/06/2003
Modifié le ... :   /  /
Description .. : Sur le changement d'année réinitialise la grille Paie
Mots clefs ... :
*****************************************************************}
procedure TFDPTableauBordGeneral.ListeAnneeChange(Sender: TObject);
var Colonne,Ligne : Integer;
begin
 inherited;
 //--- Efface le contenu de la grille Paie
 for Colonne:= 1 to GrillePaie.ColCount-1 do
  for Ligne:= 1 to 5 do
   GrillePaie.Cells [Colonne,Ligne]:='';

 InitialiserCelluleGrillePaie (Sender);
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 23/06/2003
Modifié le ... :   /  /
Description .. : Déplace la souris sur le label "Solde Client"
Mots clefs ... :
*****************************************************************}
procedure TFDPTableauBordGeneral.LSoldeMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
 inherited;
 LSolde.Color := clBtnShadow;
end;

{***********A.G.L.***********************************************
Auteur  ...... : David CATALA
Créé le ...... : 23/06/2003
Modifié le ... :   /  /
Description .. : Déplace la souris sur le panel
Mots clefs ... :
*****************************************************************}
procedure TFDPTableauBordGeneral.Panel1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
 inherited;
 LSolde.Color := clBtnFace;
end;

//-----------------------------------------
//--- Nom   : CumulAssistantClick
//--- Objet : Affiche la fiche Assistant
//-----------------------------------------
procedure TFDPTableauBordGeneral.CumulAssistantClick(Sender: TObject);
begin
 inherited;
 // if (GrilleGI.Row>0) and (GrilleGI.Row<=GrilleGI.RowCount-1) then
 if (GrilleGI.Cells [1,GrilleGI.Row]<>'Responsable') and (GrilleGI.Objects [1,GrilleGI.Row]<>Nil)then
  begin
{$IFDEF BUREAU}   // 3OB 27/08/07 
   Aff_AssistantMission (TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).CodeAffaire,TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers,RB_MISSIONSENCOURS.Checked)
{$ENDIF}   // 3OB 27/08/07 
  end;
end;

//-----------------------------------------
//--- Nom   : DetailRealiseClick
//--- Objet : Affiche le detail realisé
//-----------------------------------------
procedure TFDPTableauBordGeneral.DetailRealiseClick(Sender: TObject);
var Affaire0,Affaire1,Affaire2,Affaire3 : String;
    Affaire,Avenant,Tiers : String;
    Ddate,Fdate:TDateTime;
begin
 inherited;
 Affaire:='';
 Ddate:=StrToDate ('01/01/1900');
 FDate:=StrToDate ('31/12/2099');

 // if (GrilleGI.Row>0) and (GrilleGI.Row<=GrilleGI.RowCount-1) then
 if (GrilleGI.Cells [1,GrilleGI.Row]<>'Responsable') and (GrilleGI.Objects [1,GrilleGI.Row]<>Nil) then
  begin
   if TInfoComplementaire(GrilleGI.Objects [1,GrilleGI.Row]).CodeAffaire<>'###TOTAUX###' then
    begin
     Affaire0:='A';
     Affaire1:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire1;
     Affaire2:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire2;
     Affaire3:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire3;
     Avenant:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Avenant;
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;

     if Affaire1 <> '' then
      begin
       Affaire:=CodeAffaireRegroupe(Affaire0,Affaire1,Affaire2,Affaire3,Avenant,taModif,false,false,False);
       if not ExisteAffaire (Affaire,'') then
        begin
         if TeststCleAffaire(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers,False,False,False,True) <> 1 then
          Affaire := '';
        end;
      end;

     if (Affaire<>'') then
      AGLLanceFiche ('AFF','AFACTIVITECON_MUL','','','AFF='+affaire+ ';DATEDEB='+ DateToStr(Ddate)+ ';DATEFIN='+ DateToStr(FDate))
    end
   else
    begin
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;

     if (RB_MISSIONSENCOURS.Checked) then
      AGLLanceFiche ('AFF','AFACTIVITECON_MUL','','','CLIENT='+Tiers+';ETATAFF=ENC;DATEDEB='+ DateToStr(Ddate)+ ';DATEFIN='+ DateToStr(FDate))
     else
      AGLLanceFiche ('AFF','AFACTIVITECON_MUL','','','CLIENT='+Tiers+';DATEDEB='+ DateToStr(Ddate)+ ';DATEFIN='+ DateToStr(FDate))
    end;
  end;
end;

//-----------------------------------------
//--- Nom   : DetailRealiseClick
//--- Objet : Affiche le detail realisé
//-----------------------------------------
procedure TFDPTableauBordGeneral.DetailFactureClick(Sender: TObject);
var Affaire0,Affaire1,Affaire2,Affaire3 : String;
    Affaire,Avenant,Tiers : String;
    Ddate,Fdate:TDateTime;
begin
 inherited;

 Affaire:='';
 Ddate:=StrToDate ('01/01/1900');
 FDate:=StrToDate ('31/12/2099');

 // if (GrilleGI.Row>0) and (GrilleGI.Row<=GrilleGI.RowCount-1) then
 if (GrilleGI.Cells [1,GrilleGI.Row]<>'Responsable') and (GrilleGI.Objects [1,GrilleGI.Row]<>Nil) then
  begin
   if TInfoComplementaire(GrilleGI.Objects [1,GrilleGI.Row]).CodeAffaire<>'###TOTAUX###' then
    begin
     Affaire0:='A';
     Affaire1:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire1;
     Affaire2:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire2;
     Affaire3:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire3;
     Avenant:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Avenant;
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;
     if Affaire1 <> '' then
      begin
       Affaire:=CodeAffaireRegroupe(Affaire0,Affaire1,Affaire2,Affaire3,Avenant,taModif,false,false,False);
       if not ExisteAffaire (Affaire,'') then
        begin
         if TeststCleAffaire(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers,False,False,False,True) <> 1 then
          Affaire := '';
        end;
      end;

     if (Affaire<>'') then
       AGLLanceFiche ('AFF','AFREGRLIGNE_MUL','GL_TIERS='+Tiers+';GL_AFFAIRE='+Affaire,'','NOAFFAIRE;NATUREAUXI:CLI;NATURE=BUR;TABLE:GL;DATEDEB:'+datetostr(Ddate)+';DATEFIN:'+datetostr(Fdate)+';AFFAIRE:'+affaire)
    end
   else
    begin
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;
     // #### pour la ligne de totaux, il faut distinguer selon qu'on a "missions en cours" ou "toutes les missions"
     // => cela fonctionne pour toutes, il manque de filtrer sur ENC derrière le code mission pour les encours
     AGLLanceFiche ('AFF','AFREGRLIGNE_MUL','GL_TIERS='+Tiers,'','NOAFFAIRE;NATUREAUXI:CLI;NATURE=BUR;TABLE:GL;DATEDEB:'+datetostr(Ddate)+';DATEFIN:'+datetostr(Fdate))
    end;
  end;
end;

//------------------------------------------------------------------------------
//--- Nom   : MnGrandLivreClick
//--- Objet : JP 26/08/05 - grand livre GI sur la mission et le tiers spécifié
//------------------------------------------------------------------------------
procedure TFDPTableauBordGeneral.mnGrandLivreClick(Sender: TObject);
var Affaire0,Affaire1,Affaire2,Affaire3 : String;
    Affaire,Avenant,Tiers : String;
begin
 if (AGLJaiLeDroitFiche (['AFGRANDLIVRE','ACTION=CONSULTATION'], 2)) and (GrilleGI.Cells [1,GrilleGI.Row]<>'Responsable') and (GrilleGI.Objects [1,GrilleGI.Row]<>Nil) then
  begin
   if TInfoComplementaire(GrilleGI.Objects [1,GrilleGI.Row]).CodeAffaire<>'###TOTAUX###' then
    begin
     Affaire0:='A';
     Affaire1:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire1;
     Affaire2:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire2;
     Affaire3:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Affaire3;
     Avenant:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Avenant;
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;

     Affaire:=CodeAffaireRegroupe(Affaire0,Affaire1,Affaire2,Affaire3,Avenant,taModif,false,false,False);
     if not ExisteAffaire (Affaire,'') then
      begin
       if TeststCleAffaire(Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant,Tiers,False,False,False,True) <> 1 then
        Affaire := '';
      end;

     if (Affaire<>'') then
      begin
       AFLanceFiche_GrandLivre ('NOFILTRE:NOFILTRE;TYPETB:AFF;AFFAIRE:'+Affaire+';TIERS:'+Tiers);
       InitialiserTableGI; // FQ11638 - car le grand livre recalcule AFTABLEAUBORD !
      end
    end
   else
    begin
     Tiers:=TInfoComplementaire (GrilleGI.Objects [1,GrilleGI.Row]).Tiers;
     if (RB_MISSIONSENCOURS.Checked) then
      begin
       AFLanceFiche_GrandLivre ('NOFILTRE:NOFILTRE;TYPETB:AFF;ETATAFF:ENC;TIERS:'+Tiers);
       InitialiserTableGI; // FQ11638 - car le grand livre recalcule AFTABLEAUBORD !
      end
     else
      AFLanceFiche_GrandLivre ('NOFILTRE:NOFILTRE;TYPETB:AFF;TIERS:'+Tiers);
    end;
  end;
end;

//-----------------------------------------
//--- Nom   : PurgerComptaClick
//--- Objet : Purge la table DPTABCOMPTA
//-----------------------------------------
procedure TFDPTableauBordGeneral.PurgerComptaClick(Sender: TObject);
var Msg : String;
begin
 inherited;
 if ExisteSql ('Select 1 from DPTABCOMPTA where DTC_NoDossier="'+NumDossier+'"') then
  begin
   Msg:='Voulez-vous supprimer les données comptable.';
   if (PgiAsk (Msg,TitreHalley)=MrYes) then
    begin
     ExecuteSQl ('Delete from DPTABCOMPTA where DTC_NoDossier="'+NumDossier+'"');
     InitialiserLibelleGrilleCompta (Sender);
    end;
  end;
end;

//------------------------------------------
//--- Nom   : PurgerPaieClick
//--- Objet : Purge la table DPTABGENPAIE
//------------------------------------------
procedure TFDPTableauBordGeneral.PurgerPaieClick(Sender: TObject);
var Colonne,Ligne : Integer;
    Msg           : String;
begin
 inherited;
 if ExisteSql ('Select DT1_ANNEE from DPTABGENPAIE where DT1_NODOSSIER="'+NumDossier+'"') then
  begin
   Msg:='Voulez-vous supprimer les données paie ?';
   if (PgiAsk (Msg,TitreHalley)=MrYes) then
    begin
     ExecuteSQl ('Delete from DPTABGENPAIE where DT1_NoDossier="'+NumDossier+'"');
     //--- Efface le contenu de la grille Paie
     for Colonne:= 1 to GrillePaie.ColCount-1 do
      for Ligne:= 1 to 5 do
       GrillePaie.Cells [Colonne,Ligne]:='';
     ListeAnnee.Clear;
     InitialiserCelluleGrillePaie (Sender);
    end;
  end;
end;

//-------------------------------------------
//--- Nom   : PurgerGIClick
//--- Objet : Purge la table AFTABLEAUBORD
//-------------------------------------------
procedure TFDPTableauBordGeneral.PurgerGIClick(Sender: TObject);
var Msg : String;
begin
 inherited;
 if ExisteSql ('Select 1 from AFTABLEAUBORD') then
  begin
   Msg:='Voulez-vous supprimer les données GI ?';
   if (PgiAsk (Msg,TitreHalley)=MrYes) then
    begin
     ExecuteSQl ('Delete from AFTABLEAUBORD');
     InitialiserLibelleGrilleGI (Sender);
    end;
  end;
end;


procedure TFDPTableauBordGeneral.RB_MISSIONSTOUTESClick(Sender: TObject);
begin
 inherited;
 RB_MISSIONSENCOURS.Checked:=False;
 MoveCurProgressForm('Gestion Interne');
 InitialiserCelluleGrilleGI (Sender);
 FiniMoveProgressForm;
end;

procedure TFDPTableauBordGeneral.RB_MISSIONSENCOURSClick(Sender: TObject);
begin
 inherited;
 RB_MISSIONSTOUTES.Checked:=False;
 MoveCurProgressForm('Gestion Interne');
 InitialiserCelluleGrilleGI (Sender);
 FiniMoveProgressForm;
end;

function TFDPTableauBordGeneral.IsDossierAgricole : Boolean;
begin
 Result:=ExisteSQl('select DFI_REGLEFISC from DPFISCAL ' +
                   'where DFI_REGLEFISC="BA" and DFI_GUIDPER="'+GetGuidPer(NumDossier)+'"');
 //LM20070412 ExisteSQl('SELECT DOR_REGLEFISC FROM DPORGA Where DOR_REGLEFISC="BA" and DOR_GUIDPER="'+GetGuidPer(NumDossier)+'"');
end;

//------------------------------------------------------------------------------
//--- Nom   : BImprimerClick
//------------------------------------------------------------------------------
procedure TFDPTableauBordGeneral.BImprimerClick(Sender: TObject);
var UneTob : Tob;
    Indice : Integer;
begin
 TobImprimer:=TOB.Create('',nil,-1);

 //--- Récupération des informations Critères
 UneTob:=DonnerInformationCritere;
 if (UneTob.Detail.Count>0) then
  ConstruireTobImprimer ('ENTETE',3,UneTob.Detail [0]);

 //--- Récupération des informations concernant la grille Compta
 UneTob:=DonnerInformationGrille (GrilleCompta);
 if (UneTob.Detail.Count>0) then
  begin
   ConstruireTobImprimer ('ENTETECOMPTA',GrilleCompta.ColCount,UneTob.Detail [0]);
   for Indice:=1 to UneTob.Detail.count-1 do
    ConstruireTobImprimer ('DETAILCOMPTA',GrilleCompta.ColCount,UneTob.Detail [Indice]);
  end;
 FreeAndNil (uneTob);

 //--- Récupération des informations concernant la grille Paie
 UneTob:=DonnerInformationGrille (GrillePaie);
 if (UneTob.Detail.Count>0) then
  begin
   ConstruireTobImprimer ('ENTETEPAIE',GrillePaie.ColCount,UneTob.Detail [0]);
   for Indice:=1 to UneTob.Detail.count-1 do
    ConstruireTobImprimer ('DETAILPAIE',GrillePaie.ColCount,UneTob.Detail [Indice]);
  end;
  FreeAndNil (uneTob);

 //--- Récupération des informations concernant la grille GI
 UneTob:=DonnerInformationGrille (GrilleGI);
 if (UneTob.Detail.Count>0) then
  begin
   ConstruireTobImprimer ('ENTETEGI',GrilleGI.ColCount,UneTob.Detail [0]);
   for Indice:=1 to UneTob.Detail.count-1 do
    ConstruireTobImprimer ('DETAILGI',GrilleGI.ColCount,UneTob.Detail [Indice]);
  end;
  FreeAndNil (uneTob);

 //--- Impression de l'état Tableau de bord
 if TobImprimer.Detail.Count > 0 then
  LanceEtatTob( 'E','DPG','TB1',TobImprimer,True,False,False,nil,'','Edition Tableau de bord',False)
 else
  PgiInfo ('Aucune donnée à imprimer.',TitreHalley);
 FreeAndNil (TobImprimer);
end;

//------------------------------------------------------------------------------
//--- Nom   : DonnerInformationGrille
//------------------------------------------------------------------------------
function TFDPTableauBordGeneral.DonnerInformationCritere : Tob;
Var TobCritere,UneTobEnreg : Tob;
begin
 TobCritere:=Tob.Create ('',nil,-1);
 UneTobEnreg:=TOB.Create ('',TobCritere,-1);

 UneTobEnreg.AddChampSupValeur ('CHP00',ListeAnnee.Text);
 if (Rb_MissionsEncours.Checked) then
  UneTobEnreg.AddChampSupValeur ('CHP01','Uniquement les missions en cours')
 else
  UneTobEnreg.AddChampSupValeur ('CHP01','Toutes les missions');
 UneTobEnreg.AddChampSupValeur ('CHP02',Solde.Caption);

 Result:=TobCritere;
end;

//------------------------------------------------------------------------------
//--- Nom   : DonnerInformationGrille
//------------------------------------------------------------------------------
function TFDPTableauBordGeneral.DonnerInformationGrille (UneGrille : THGrid) : Tob;
Var TobGrille,UneTobEnreg : Tob;
    NomChp,ValChp         : String;
    IndiceCol,IndiceLigne : Integer;
begin
 TobGrille:=Tob.Create ('',nil,-1);

 for IndiceLigne:=0 to UneGrille.RowCount-1 do
  begin
   UneTobEnreg:=TOB.Create ('',TobGrille,-1);
   for IndiceCol:=0 to UneGrille.ColCount-1 do
    begin
     NomChp:='CHP'+Format ('%2.2d',[IndiceCol]);
     ValChp:=UneGrille.Cells [IndiceCol,IndiceLigne];
     UneTobEnreg.AddChampSupValeur (NomChp,ValChp);
    end;
  end;

 Result:=TobGrille;
end;

//------------------------------------------------------------------------------
//--- Nom   : ConstruireTobImprimer
//------------------------------------------------------------------------------
procedure TFDPTableauBordGeneral.ConstruireTobImprimer (UnLibelleLigne : String; NbreChp : Integer; UneTob : Tob);
var Indice      : Integer;
    TbChpVal    : String;
    UneTobEnreg : Tob;
begin

 UneTobEnreg:=TOB.Create ('',TobImprimer,-1);
 for Indice:=0 to 14 do
  UneTobEnreg.AddChampSupValeur ('TBChp'+intToStr (Indice),'');

 UneTobEnreg.AddChampSupValeur ('TBChp0',UnLibelleLigne);
 for Indice:=1 to NbreChp do
  begin
   TbChpVal:=UneTob.Getvalue ('CHP'+Format ('%2.2d',[Indice-1]));
   if (TbChpVal<>'') then
    UneTobEnreg.PutValue ('TBChp'+IntToStr (Indice),TbChpVal)
   else
    UneTobEnreg.PutValue ('TBChp'+IntToStr (Indice),' ');
  end;
end;

end.

