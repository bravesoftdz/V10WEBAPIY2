{***********UNITE*************************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Consultation du contenu d'un fichier TOX
Mots clefs ... : TOX; DATA
*****************************************************************}
unit PaieToxConsultBis;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  HTB97,
  ComCtrls,
  StdCtrls,
  ExtCtrls,
  HPanel,
  uTob,
  hmsgbox,
  Grids,
  Hctrls,
  DBTables,
  PaieToxMoteur,
  HSysMenu,
  uToxClasses,
  uToxConst,
  HEnt1,
  UIUtil,
  utoz ;

type
  TConsultMessageToxBis = class(TForm)
    HPanel1: THPanel;
    HPanel3: THPanel;
    OD: TOpenDialog;
    TreeTOX: TTreeView;
    Tableau: THGrid;
    Panel1: TPanel;
    FICHIER: TEdit;
    NBENREG: TEdit;
    NBERREUR: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Panel2: TPanel;
    MESSINTEGRATION: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    NOMTABLE: TEdit;
    Panel3: TPanel;
    Bopen: TToolbarButton97;
    OuvertureFichier: TLabel;
    DonneeRejetee: TCheckBox;
    DonneeIntegree: TCheckBox;
    Label7: TLabel;
    NBTraitement: TEdit;
    HMTRAD: THSystemMenu;
    Label8: TLabel;
    Label9: TLabel;
    SiteEmet: TEdit;
    SiteDest: TEdit;
    Label10: TLabel;
    DateEmis: TEdit;
    DeviseEmis: TEdit;
    Label3: TLabel;
    Label11: TLabel;
    DeviseContre: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure BopenClick(Sender: TObject);
    procedure BcancelClick(Sender: TObject);
    function  ChargeTOX (TT : TOB) : Integer ;
    procedure TreeTOXChange(Sender: TObject; Node: TTreeNode);
    procedure OnNeFaitRien (Sender: TObject; Node: TTreeNode);
    procedure TreeTOXCustomDrawItem(Sender: TCustomTreeView;
                                    Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure ModifieAffichage ( Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
                                var DefaultDraw: Boolean);
    procedure ChangeInfoEcran (Sender: TObject);
    function  AfficheInfoTechnique (T : TOB) : integer;
    procedure FormCreate(Sender: TObject);
    procedure AfficheInfoFichier;
    procedure AfficheLaTox;
    procedure AfficheLaToxTech;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    ToxFichier  : TOB ;      // TOB du fichier TOX
    ToxMessage  : TOB ;      // TOB des messages d'intégration
    TobInfoTech : TOB ;      // TOB des informations techniques
    NameTable   : string ;   // Nom de la table courante
    FileName    : string ;   // Nom du fichier TOX
    Compress    : boolean ;
  end;


////////////////////////////////////////////////////////////////////////////////
// Procédures ou fonctions à exporter
////////////////////////////////////////////////////////////////////////////////
  procedure PaieConsultManuelTox ;
  procedure PaieConsulteAutoTox ( FileName : string; MessageDonnee: boolean);

implementation

{$R *.DFM}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Création de la forme et ouverture du fichier TOX passé en
Suite ........ : paramètres
Suite ........ : Fonction appelée à partir de la visualisation des corbeilles
Mots clefs ... : TOX;DATA
*****************************************************************}
procedure PaieConsulteAutoTox ( FileName : string; MessageDonnee: boolean);
var
  X: TConsultMessageToxBis;
begin
  X := TConsultMessageToxBis.Create ( Application ) ;
  X.FileName := FileName ;
  X.OuvertureFichier.Visible := False ;
  X.Bopen.visible := False;
  SourisSablier ;
  if MessageDonnee then X.AfficheLaTox
  else begin
    X.AfficheLaToxTech ;
  end;
  SourisNormale ;
  X.ShowModal ;
  X.Free ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Création de la forme
Suite ........ : Sélection manuelle d'un fichier TOX pour afficher son 
Suite ........ : contenu
Mots clefs ... : TOX;DATA
*****************************************************************}
procedure PaieConsultManuelTox ( ) ;
var
  X  : TConsultMessageToxBis;
  PP : THPanel ;
begin
  PP:=FindInsidePanel ;
  X := TConsultMessageToxBis.Create ( Application ) ;
  if PP=Nil then
  begin
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   end else
   begin
     InitInside(X,PP) ;
     X.Show ;
   end ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Charge le fichier TOX dans la TOB "ToxFichier"
Mots clefs ... : TOX
*****************************************************************}
function TConsultMessageToxBis.ChargeTOX(TT : TOB) : Integer ;
BEGIN
  ToxFichier := TT ;
  result     := 0 ;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Fermeture de la Forme
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Fermeture de la Forme
Suite ........ :   -> Libération des TOB
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.FormDestroy(Sender: TObject);
begin
  if ToxFichier   <> nil then Toxfichier.Free ;
  if ToxMessage   <> nil then ToxMessage.Free ;
  if TobInfoTech  <> nil then TobInfoTech.Free ;
  //
  // Si fichier zippé, on le supprime
  //
  if Compress then
  begin
    DeleteFile ( FileName ) ;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Chargement du fichier TOX en TOB
Suite ........ : Chargement des messages en TOB
Suite ........ : Affichages
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.AfficheLaTox;
var NbEnr, NbErr : integer;
begin
  //
  // Chargement du fichier en TOB
  //
  if ToxFichier <> nil then ToxFichier.free;
  ToxFichier := Tob.Create ('THE TOX', Nil, -1 );
  TOBLoadFromBinFile (FileName, ChargeTOX, nil);
  //
  // Création et constitution de la TOb des messages
  //
  if ToxMessage <> nil then ToxMessage.free;
  ToxMessage := Tob.Create ('THE MESS', Nil, -1 );
  ToxMessage.AddChampSup ('INFORMATIONS', False);
  ToxMessage.PutValue    ('INFORMATIONS', 'INFORMATIONS');
  //
  // Affichage du nom du fichier , du nombre de traitements
  //
  AfficheInfoFichier;
  //
  // Consultation des messages de la TOX
  //
  PaieChargeMessageDeLaTox (ToxFichier, ToxMessage, DonneeIntegree.checked, DonneeRejetee.checked, NbEnr, NbErr);
  //
  // Affichage du nbre d'enreg et du nombre d'erreurs.
  //
  NbEnreg.Text  := IntToStr (NbEnr) ;
  NbErreur.Text := IntToStr (NbErr) ;

  if ToxMessage.Detail.Count <= 0 then ToxMessage.ClearDetail;
{$IFNDEF EAGLSERVER}
  ToxMessage.PutTreeView ( TreeTOX, Nil, 'INFORMATIONS;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE');
{$ENDIF}  
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. :  Chargement du fichier TOX en TOB
Suite ........ :  Chargement des messages en TOB
Suite ........ :  Affichages
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.AfficheLaToxTech;
var NbEnr, NbErr : integer;
begin
  NbEnr := 0;
  NbErr := 0;
  //
  // Redéfinition de la fonction OnChange du TreeTOX pour ne rien faire
  //
  TreeTox.OnChange:=OnNeFaitRien;
  TreeTox.OnCustomDrawItem:=ModifieAffichage;
  DonneeIntegree.Enabled := False;
  DonneeRejetee.Enabled  := False;
  //
  // Chargement du fichier en TOB
  //
  if ToxFichier <> nil then ToxFichier.free;
  ToxFichier := Tob.Create ('THE TOX', Nil, -1 );
  TOBLoadFromBinFile (FileName, ChargeTOX, nil);
  //
  // Affichage du nom du fichier , du nombre de traitements
  //
  AfficheInfoFichier;
  //
  // Consultation des messages de la TOX
  //
  //ChargeMessageDeLaTox (ToxFichier, ToxMessage, DonneeIntegree.checked, DonneeRejetee.checked, NbEnr, NbErr);
  //
  // Affichage du nbre d'enreg et du nombre d'erreurs.
  //
  NbEnreg.Text  := IntToStr (NbEnr) ;
  NbErreur.Text := IntToStr (NbErr) ;

  if ToxFichier.Detail.Count < 0 then ToxFichier.ClearDetail;
{$IFNDEF EAGLSERVER}
  ToxFichier.PutTreeView ( TreeTOX, Nil, 'EVENEMENT;CONDITION;TOX_TABLE;DONNEES;TOX_TABLE;DONNEE');
{$ENDIF}  
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Sélection manuelle du fichier du fichier TOX
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.BopenClick(Sender: TObject);
var MaToz        : Toz     ;
    Okay         : boolean ;
begin
  MaToz := nil ;

  if OD.Execute then
  begin
    Compress := False ;

    if UpperCase ( ExtractFileExt ( OD.FileName ) ) = '.ZIP' then
    begin
      Okay := False ;
      //
      // Extraction du fichier .TOX du .ZIP
      TRY
        TRY
          MaToz := TOZ.Create ;
          if MaToz.OpenZipFile ( OD.FileName, moOpen ) then
          begin
            if MaToz.OpenSession (osExt) then
            begin
              if MaToz.SetDirOut (ExtractFilePath ( OD.FileName )) then
              begin
                if MaToz.CloseSession then
                begin
                  // Suppression de la TOZ
                  Okay := True ;
                  MaToz.Free ;
                  MaToz := Nil ;
                  //
                  // Ré-initialise le fichier chargé
                  //
                  FileName := Copy (OD.FileName , 1, Length ( OD.FileName) -4 ) ;
                  Compress := True ;
                end ;
              end ;
            end ;
          end ;
        EXCEPT
          on Erreur : ETozErreur do
          begin
            Okay := False ;
          end ;
        end;
      FINALLY
        if MaToz <> Nil then MaToz.Free ;
      end;
    end else
    begin
      Okay := True ;
      FileName := OD.FileName ;
    end;

    if Okay then AfficheLaTox ;

  end;
end;

procedure TConsultMessageToxBis.BcancelClick(Sender: TObject);
begin
  Close ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Affiche dans un THGRID le contenu de l'enregistrement 
Suite ........ : TOX
Suite ........ : 
Suite ........ :          1 - Création de la TOB virtuelle mère TobInfoTech
Suite ........ :          2 - Création de la TOB fille rattachée à la table
Suite ........ :          3 - Affichage
Mots clefs ... : TOX
*****************************************************************}
function TConsultMessageToxBis.AfficheInfoTechnique (T : TOB) : integer;
var TobFille : TOB;
begin
  if TobInfoTech <> nil then TobInfoTech.free ;
  TobInfoTech := TOB.Create('INFO TECHN', Nil, -1);

  TobFille := TOB.Create (NameTable, TobInfoTech, -1);
  TobFille.Dupliquer (T, False, True, True);

  if Tableau.Visible = False then Tableau.Visible := True;
  Tableau.VidePile ( False ) ;
  Tableau.ColCount := 2 ;

  TobInfoTech.PutGridDetail(Tableau, True, False, '', True);
  result := 0;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Récupération et affichage des informations contenues dans
Suite ........ : l'enveloppe du fichier
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.AfficheInfoFichier ;
var
    PEnveloppes: TCollectionEnveloppes ;
    FileNameEnv : string        ;
    length      : integer       ;
begin
  //
  // Nom du fichier et nombre de traitement
  //
  Fichier.text         := FileName;
  MessIntegration.Text := '';
  if not ToxFichier.FieldExists ('NB_TRAITEMENT') then NbTraitement.text := '0'
  else NbTraitement.Text := IntToStr (ToxFichier.GetValue ('NB_TRAITEMENT'));
  //
  // Récupération des informations contenues dans l'enveloppe
  //
  trim (FileName);
  length      := strlen (PChar (FileName));
  FileNameEnv := copy (FileName, 0, length-4) + '.ENV' ;

  PEnveloppes:=Nil ;

  TRY

  Penveloppes:=TCollectionEnveloppes.Create(TCollectionEnveloppe) ;
  PEnveloppes.LoadEnveloppe(FileNameEnv) ;

    //
    // Recherche dans l'enveloppe des sites émetteur / destinataire + Date d'émission
    //
    with PEnveloppes.Items[0] do
    begin
      SiteEmet.Text := eLibelle ; // PEnv^.Emetteur^.Libelle;
      SiteDest.Text := dLibelle ; // PEnv^.Destinataire^.Libelle;
      DateEmis.text := FormatDateTime('dd/mm/yyyy à hh:nn:ss',DateMsg); // + ' à ' +TimeToStr(DateMsg);
  END;    
  FINALLY
    //
    // Libération de l'enveloppe
    //
    if PEnveloppes<>Nil then PEnveloppes.Free ;
  END ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. :  Sélection d'une autre ligne dans le TreeView
Suite ........ :    -> MAJ des zones d'affichage (message d'erreur, table)
Suite ........ :    -> Recherche de l'info technique associée, et affichage
Suite ........ : dans le THGRID
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.TreeTOXChange(Sender: TObject; Node: TTreeNode);
var TOBB          : TOB;
    IndiceMessage : integer ;
begin
  if Node.Data <> Nil then
  begin
     TOBB := TOB ( Node.Data ) ;

     if TOBB.FieldExists ('MESS_ERREUR') then
     begin
       if (TOBB.FieldExists ('ERREUR')) and (StrToInt(TOBB.GetValue ('ERREUR')) > 0) then MessIntegration.Font.color := clRed
       else MessIntegration.Font.Color := ClBlack;
       MessIntegration.Text := TOBB.GetValue ('MESS_ERREUR');
     end;
     if TOBB.FieldExists ('TABLE')       then NameTable            := TOBB.GetValue ('TABLE');
     NomTable.Text := NameTable ;

     if TOBB.FieldExists ('INDICE') then IndiceMessage := TOBB.GetValue ('INDICE')
     else IndiceMessage := 0 ;

     if (IndiceMessage > 0) then
     begin
       ToxFichier.ParcoursTraitement (['TOX_INDICE'], [IndiceMessage], True, AfficheInfoTechnique);
     end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Sélection d'une autre ligne dans le TreeView
Suite ........ :   -> Mode "Affichage Technique" : on ne fait rien
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.OnNeFaitRien(Sender: TObject; Node: TTreeNode);
begin
//
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Permet de mettre en rouge les lignes messages des 
Suite ........ : enregistrements en erreur dans le TreeView
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.ModifieAffichage (Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
                                                   var DefaultDraw: Boolean);
var TOBB    : TOB ;
    Libelle : string ;
begin

  if Node.Data <> Nil then
  begin
    TOBB := Tob (Node.data);
    if TOBB.FieldExists ('EVENEMENT') then
    begin
      Libelle := Node.text;
      Node.Text := 'Evènement ' + Libelle ;
    end
    else if TOBB.FieldExists ('CONDITION') then
    begin
      Libelle := Node.text;
      Node.Text := 'Requètes ' + Libelle ;
    end;
    //else Node.Text := 'Données ......';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Permet de mettre en rouge les lignes messages des 
Suite ........ : enregistrements en erreur dans le TreeView
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.TreeTOXCustomDrawItem(
  Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
  var DefaultDraw: Boolean);
var TOBB : TOB ;
begin

  if Node.Data <> Nil then
  begin
    TOBB := Tob (Node.data);
    If (TOBB.FieldExists ('ERREUR')) and (TOBB.GetValue ('ERREUR') > 0) then
    begin
      Sender.Canvas.Font.color := ClRed;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Changement des informations à afficher : données 
Suite ........ : intégrées, données rejetées
Suite ........ : 
Suite ........ : -> On vide tout, on recharge tout et on ré-affiche tout
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.ChangeInfoEcran (Sender: TObject);
var Nbenr, Nberr : integer;
begin
  if ToxFichier <> nil then
  begin
    //
    // Remise à 0 de la TOB des messages
    //
    if ToxMessage <> nil then ToxMessage.free;
    //
    // Création de la TOB des messages
    //
    ToxMessage := Tob.Create ('THE MESS', Nil, -1 );
    ToxMessage.AddChampSup   ('INFORMATIONS', False);
    ToxMessage.PutValue      ('INFORMATIONS', 'INFORMATIONS');
    //
    // Initialisation des zones d'affichage
    //
    MessIntegration.Text := '';
    //
    // Vidage du GRID des infos techniques
    //
    Tableau.VidePile ( False ) ;
    Tableau.Visible := False ;
    //
    // Chargement messages de la TOX
    //
    PaieChargeMessageDeLaTox (ToxFichier, ToxMessage, DonneeIntegree.checked, DonneeRejetee.checked, NbEnr, NbErr);
    //
    // Affichage si données
    //
    if ToxMessage.Detail.Count <= 0 then ToxMessage.ClearDetail;
{$IFNDEF EAGLSERVER}
    ToxMessage.PutTreeView ( TreeTOX, Nil, 'INFORMATIONS;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE;MESS_DONNEE')
{$ENDIF}    
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Permet de sélectionner uniquement les fichiers TOX
Mots clefs ... : TOX
*****************************************************************}
procedure TConsultMessageToxBis.FormCreate(Sender: TObject);
begin
  OD.Filter := 'Fichiers TOX|*.TOX*';
  TreeTOX.HideSelection := False ;
  HMtrad.ResizeGridColumns (Tableau);
end;

end.
