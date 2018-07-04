unit dpPublicationCD;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ComCtrls, HSysMenu, hmsgbox, StdCtrls, ExtCtrls, HPanel, HTB97,
  Hctrls, Menus, Buttons,
  contnrs, uTOB,
  cdBurn, jpeg // Pour gravage CD (encapsulation IMAPI)
  ;

const
  ccConsulterDocumentAutre     = 187310 ;

type
  TFPublicationCD = class(TFAssist)
    TSDossierPermanent: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    PopStructure: TPopupMenu;
    TabSheet4: TTabSheet;
    Label1: TLabel;
    Label8: TLabel;
    EResume: TRichEdit;
    MPublier: TMenuItem;
    MRegroupe: TMenuItem;
    MMoisAnnee: TMenuItem;
    MNature: TMenuItem;
    MAuteur: TMenuItem;
    PStructBoutons: TPanel;
    BDefaut_DP: THSpeedButton;
    LExplic_DP: TLabel;
    Panel1: TPanel;
    MNoRegroupe: TMenuItem;
    GBOptions: TGroupBox;
    Label9: TLabel;
    EPathCible: TEdit;
    BSelectPath: THSpeedButton;
    Label7: TLabel;
    BApercuHtml: THSpeedButton;
    GBCabinet: TGroupBox;
    Label5: TLabel;
    ELogoFile: TEdit;
    BSelectLogo: THSpeedButton;
    BApercuLogo: THSpeedButton;
    Label11: TLabel;
    ESiteCab: TEdit;
    BApercuSiteCab: THSpeedButton;
    GBPerso: TGroupBox;
    Label10: TLabel;
    Label6: TLabel;
    ECommentaire: TEdit;
    ETitleHtml: TEdit;
    CBGraveurs: TComboBox;
    LGraveurs: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    EPathModele: TEdit;
    BSelectModele: THSpeedButton;
    BApercuPubli: THSpeedButton;
    MPasPublier: TMenuItem;
    Label2: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    CBJaquette: TComboBox;
    BapercuJaquette: THSpeedButton;
    ICDCLIENT: TImage;
    BImpJaquette: THSpeedButton;
    BGraveCD: THSpeedButton;
    TabSheet5: TTabSheet;
    Label15: TLabel;
    EMoisAnnee_DP: TEdit;
    EAuteur_DP: TEdit;
    Label16: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label17: TLabel;
    MVBNature_DP: THMultiValComboBox;
    CBAboPub_DP: TCheckBox;
    Label20: TLabel;
    TabSheet6: TTabSheet;
    TSDossierAnnuel: TTabSheet;
    LFiltres_DA: TLabel;
    CBAboPub_DA: TCheckBox;
    Label22: TLabel;
    EMoisAnnee_DA: TEdit;
    Label23: TLabel;
    MVBNature_DA: THMultiValComboBox;
    Label24: TLabel;
    Label25: TLabel;
    EAuteur_DA: TEdit;
    Panel2: TPanel;
    Label21: TLabel;
    Panel3: TPanel;
    BDefault_DA: THSpeedButton;
    LExplic_DA: TLabel;

    TVPublication_DP : TTreeView;
    TVPublication_DA : TTreeView;

    procedure PopStructurePopup(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TVPublication_Expanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
    procedure BDefaut_DPClick(Sender: TObject);
    procedure MPublierClick(Sender: TObject);
    procedure TVPublication_Change(Sender: TObject; Node: TTreeNode);
    procedure MMoisAnneeClick(Sender: TObject);
    procedure MNatureClick(Sender: TObject);
    procedure MAuteurClick(Sender: TObject);
    procedure MNoRegroupeClick(Sender: TObject);
    procedure BSelectLogoClick(Sender: TObject);
    procedure BApercuLogoClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure BSelectPathClick(Sender: TObject);
    procedure BApercuSiteCabClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BApercuHtmlClick(Sender: TObject);
    procedure TVPublication_DblClick(Sender: TObject);
    procedure TVPublication_KeyDown(Sender:TObject; var Key:Word; Shift:TShiftState);
    procedure BSelectModeleClick(Sender: TObject);
    procedure BApercuPubliClick(Sender: TObject);
    procedure MPasPublierClick(Sender: TObject);
    procedure BapercuJaquetteClick(Sender: TObject);
    procedure CBGraveursChange(Sender: TObject);
    procedure BGraveCDClick(Sender: TObject);
    procedure BImpJaquetteClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);

  public
        m_TOBData    : TOB;
        strPathCible : string;

        function   GetResume (iNumLigne:integer):string;

  private
         m_BurnCD      :TXPBurn;
         m_bGenerated  :boolean;
         //m_bFiltered   :boolean;

         procedure LoadJaquetteList;
         function  HasDocuments         (sCodeGed: String):boolean;

         procedure DeleteDirectory      (strPath:string);

         procedure LoadData             (strLogoFile:string);
         procedure DeleteNoPublished    (TvPublication : TTreeView; TN:TTreeNode);

         procedure WriteDocuments       (HtmlLines:TStringList; DocList:TObjectList);
         procedure CreateHtml;          //; bJaquette:boolean=FALSE);
         function  BurnCD               (strPath:string):boolean;


         procedure FreeGedNodes (TvPublication : TTreeView);
         function  AjouteNoeudDocuments (TvPublication : TTreeView; TNGed:TTreeNode):boolean;
         procedure BuildPublication (TvPublication : TTreeView; SArboGed : String);
         procedure SelectNode (TvPublication : TTreeView; Node:TTreeNode);
         procedure ChangeCDPubli (TvPublication : TTreeView; TN:TTreeNode; strCDPubli:string=#0; bOnlyDocs:boolean=TRUE; bPropagate:boolean=TRUE);
         procedure ChangeReg (TvPublication : TTreeView; StrNewReg : string);
         procedure FilterSelection (TvPublication : TTreeView;  SArboGed : String);
         procedure ConstruireHTMLPublication (TN : TTreeNode; HTMLNewLines : TStringList; DocList : TObjectList; bResume : Boolean);
  end;

const
     // Recopié de UGedDp :
     // SmallImgList dans menudisp
     cImgDossierRougeSablier = 0;
     cImgCommentaire         = 1;
     cImgBureau              = 2;
     cImgDocuments           = 3;
     cImgDocumentAbsent      = 4;
     cImgDocumentInconnu     = 5;
     cImgFiche               = 6;
     cImgListe               = 7;
     cImgDocument            = 8;
     cImgArmoireFerme        = 9;
     cImgArmoireOuverte      = 10;
     cImgClasseurFerme       = 11;
     cImgClasseurOuvert      = 12;
     cImgIntercalFerme       = 11;
     cImgIntercalOuvert      = 12;
     cImgDossierClient       = 13;
     cImgCDPubli             = 14;
     cImgCDPubliInactif      = 18;
     cImgCDPubliNot          = 19;
     cImgInternetExplorer    = 22;

procedure StartPublicationCD;

var
  FPublicationCD: TFPublicationCD;

implementation

uses utilGED, YNewDocument, HdtLinks, uDossierSelect, galMenuDisp,
     shellapi, paramsoc, hent1, filectrl,
     AnnOutils, // $$$ JP 16/05/06: pour JaiLeDroitConceptBureau
{$IFNDEF EAGLCLIENT}
     edtREtat,
{$ELSE}
     UtilEagl,
{$ENDIF}
     edtEtat;


{$R *.DFM}

procedure StartPublicationCD;
begin
     FPublicationCD := TFPublicationCD.Create (Application);
     try
        FPublicationCD.ShowModal;
     finally
            FreeAndNil (FPublicationCD);
     end;
end;

procedure TFPublicationCD.FormCreate (Sender: TObject);
var
   TOBCab    :TOB;
   i         :integer;
begin
     inherited;

     // Construction de l'arborescence en fonction du paramétrage défini: publication, regroupememnt...
     TOBCab := TOB.Create ('le cabinet', nil, -1);
     try
        SourisSablier;
        BuildPublication (TvPublication_DP,'D');
        BuildPublication (TvPublication_DA,'A');

        // On lit les "coordonnées" du cabinet (site web)
        TOBCab.LoadDetailFromSQL ('SELECT ANN_SITEWEB FROM ANNUAIRE,DOSSIER WHERE DOS_NODOSSIER="' + V_PGI.NoDossier + '" AND ANN_GUIDPER=DOS_GUIDPER');
        if TOBCab.Detail.Count > 0 then
           ESiteCab.Text := TOBCab.Detail [0].GetString ('ANN_SITEWEB');
     finally
            TOBCab.Free;
            SourisNormale;
     end;

     // Répertoire de préparation par défaut: rep' dossier + sous rep CD-GED
     if VH_DOSS.PathDos <> '' then
         EPathCible.Text := VH_DOSS.PathDos + '\CD-CLIENT';

     // Logo par défaut: celui nommé cab_logo.jpg dans le répertoire standards cabinet
     if FileExists (V_PGI.DatPath + '\cab_logo.jpg') = TRUE then
        ELogoFile.Text := V_PGI.DatPath + '\cab_logo.jpg';

     // Titre par défaut du sommaire html: fondé sur le nom dossier
     ETitleHtml.Text := 'Publication de ' + VH_Doss.LibDossier;

     // Par défaut, on sélectionne le répertoire modèle cegid
     EPathModele.Text := V_PGI.StdPath + '\Bureau\CD-CLIENT';

     // $$$ JP 03/10/06: toujours un élément: "pas de gravure"
     CBGraveurs.Items.AddObject ('<pas de gravure>', nil);
     CBGraveurs.ItemIndex := 0; // $$$ JP 03/10/06

     // On essai d'instancier le Burner de CD. Si ok, on remplit avec la liste des graveurs disponibles
     try
        m_BurnCD := TXPBurn.Create (Application);
     except
           m_BurnCD := nil;
     end;
     if m_BurnCD <> nil then
     begin
          try
             if m_BurnCD.DataRecorders.Count = 0 then
                 FreeAndNil (m_BurnCD)
             else
                 for i := 0 to m_BurnCD.DataRecorders.Count-1 do
                     CBGraveurs.Items.AddObject (m_BurnCD.DataRecorders.Items [0].DosDevice + ' ' + m_BurnCD.DataRecorders.Items [0].DisplayName, m_BurnCD.DataRecorders.Items [0]);
          except
                FreeAndNil (m_BurnCD);
          end;
     end;

     // On désactive le choix d'unité gravure si IMAPI non dispo
     if m_BurnCD = nil then
     begin
          LGraveurs.Enabled  := FALSE;
          CBGraveurs.Enabled := FALSE;
     end;

     // Chargement de la liste des états jaquette
     LoadJaquetteList;
end;

//-------------------------
//--- Nom : FreeGedNodes
//-------------------------
procedure TFPublicationCD.FreeGedNodes (TvPublication : TTreeView);
var TNChild      :TTreeNode;
begin
     TNChild := TvPublication.Items.GetFirstNode;
     while TNChild <> nil do
     begin
          TGedDPNode (TNChild.Data).Free;
          TNChild := TNChild.GetNext;
     end;
end;


//-------------------------
//--- Nom : FormDestroy
//-------------------------
procedure TFPublicationCD.FormDestroy (Sender:TObject);
begin
 FreeGedNodes (TvPublication_DP);
 FreeGedNodes (TvPublication_DA);
 FreeAndNil (m_TOBData);
 FreeAndNil (m_BurnCD);
 Inherited;
end;

// Recopier de TGedDP.HasDocuments dans ugeddp.pas
// Est-ce qu'il y a des documents dans le noeud Ged
// (Documents au sens large = y compris des commentaires)
function TFPublicationCD.HasDocuments (sCodeGed:string):boolean;
begin
     Result := ExisteSQL('SELECT DPD_CODEGED FROM DPDOCUMENT WHERE DPD_NODOSSIER="' + VH_Doss.NoDossier + '" AND DPD_CODEGED="' + sCodeGed + '"');
end;

//---------------------------------
//--- Nom : AjouteNoeudDocuments
//---------------------------------
function TFPublicationCD.AjouteNoeudDocuments (TvPublication : TTreeView; TNGed:TTreeNode):boolean;
var TN           :TTreeNode;
    TGN          :TGedDPNode;
    sCodeGed     :string;
    ParentCDPub  :string;
    niveau       :integer;
begin
     Result   := False;
     sCodeGed := '';
     Niveau   := -1;

     if (TNGed<>Nil) and (TNGed.Data<>Nil) then
     begin
          Niveau := TNGed.Level;
          with TGedDPNode (TNGed.Data) do
          begin
               sCodeGed    := CodeGed; //TGedDPNode (TNGed.Data).CodeGed;
               ParentCDPub := CDPubli;
          end;
     end;

     // A-t'on besoin du noeud Documents
     if Not HasDocuments (sCodeGed) then
        exit;

     // Niveau -1 = racine, Niveau 0 = armoires, Niveau 1 = classeurs, Niveau 2 = Intercalaire
     if niveau < 2 then
     begin
          // Création du noeud spécifique "Documents divers"
          // Les documents sous-jacents ne seront chargés que sur expanding de ce noeud.
          TGN            := TGedDPNode.Create;
          TGN.CodeGed    := sCodeGed;
          TGN.TypeGed    := 'DCS';
          TGN.SFileGUID     := '';
          TGN.SDocGUID      := '';
          TGN.TagMenu    := 0;
          TGN.IconId     := cImgDocuments;
          TGN.IconSel    := cImgDocuments;
          if niveau=-1 then
          begin
               TGN.Caption := 'Documents en attente';
               TGN.CDPubli := '';
          end
          else
          begin
               TGN.Caption := 'Documents divers';
               TGN.CDPubli := ParentCDPub;
          end;

          TN               := TvPublication.Items.AddChildObject (TNGed, TGN.Caption, TGN);
          TN.ImageIndex    := TGN.IconId;
          TN.SelectedIndex := TGN.IconSel;
          TvPublication.Items.AddChild (TN, 'Dummy');
     end
     else
         if (niveau=2) and (TGedDpNode(TNGed.Data).TypeGed='ITL') then
            // Niveau 2 = Intercalaire
            // Dans une intercalaire, il n'y a que des documents, donc on n'ajoute
            // pas un noeud intermédiaire "Documents"
            // (donc on ajoute le dummy qui permettra d'afficher les documents sur expand)
            TvPublication.Items.AddChild (TNGed, 'Dummy');

     Result := True;
end;

//-----------------------------
//--- Nom : BuildPublication
//-----------------------------
// Recopier en grande partie de TGedDP.AlimTreeViewGedDP dans ugeddp.pas
procedure TFPublicationCD.BuildPublication (TvPublication : TTreeView; SArboGed : String);
var SSQL                            : string;
    i, j, k, tmpIndex               : integer;
    TGN1, TGN2, TGN3                : TGedDPNode;
    TN1, TN2, TN3                   : TTreeNode;
    TobNiv1, TobNiv2, TobNiv3, T    : TOB;
    DataTypeLink, DataTypeLink2     : TYDataTypeLink;
    Niv1Utile, Niv2Utile, Niv3Utile : Boolean;
begin
 //--- Empêche le rafraichissement pendant le chargement
 TvPublication.Items.BeginUpdate;

 //--- On supprime les noeuds GED déjà présents
 TN1 := TvPublication.Items.GetFirstNode;
 while TN1 <> nil do
  begin
   TGedDPNode (TN1.Data).Free;
   TN1:=TN1.GetNext;
  end;

 TvPublication.Items.Clear; // Le treeview reste créé, on le réalimente à chaque chgt de dossier
 TvPublication.Images := FMenuDisp.SmallImages; // Liste des images de la publication

 //--- Niveau 1 - méthode GetTabletteSQL => ramène la requête SQL de la tablette voulue
 SSQL:='SELECT * FROM YGEDDICO AS YG1 WHERE EXISTS('
      + GetTabletteSQL ('YYGEDNIVEAU1', 'YG1.YGD_CODEGED=YGEDDICO.YGD_CODEGED', True )+ ')'
      + ' AND YGD_ARBOGED="'+SArboGed+'"'
      + ' ORDER BY YGD_TRIGED';
 TobNiv1 := Tob.Create ('Ged Niveau 1', nil, -1);
 TobNiv1.LoadDetailFromSQL (SSQL);

 try
  //--- Objet "Tablettes hiérarchiques"
  DataTypeLink := V_DataTypeLinks.FirstMaster ('YYGEDNIVEAU1', tmpIndex);
  DataTypeLink2 := V_DataTypeLinks.FirstMaster('YYGEDNIVEAU2', tmpIndex);

  for i := 0 to TobNiv1.Detail.Count-1 do
   begin
    Niv1Utile := False;
    T := TobNiv1.Detail [i];

    TGN1 := TGedDPNode.Create;
    TGN1.CodeGed := T.GetValue ('YGD_CODEGED');
    TGN1.TypeGed := T.GetValue ('YGD_TYPEGED');
    TGN1.SFileGUID := '';
    TGN1.SDocGUID := '';
    TGN1.TagMenu := T.GetValue ('YGD_TAGMENU');
    TGN1.IconId := cImgArmoireFerme;
    TGN1.IconSel := cImgArmoireFerme;
    TGN1.Caption := T.GetValue ('YGD_LIBELLEGED');
    TGN1.CDPubli := T.GetString ('YGD_CDPUBLI');

    TN1 := TvPublication.Items.AddChildObject (nil, TGN1.Caption, TGN1);
    TN1.ImageIndex := TGN1.IconId;
    TN1.SelectedIndex := TGN1.IconSel;

    //--- Niveau 2 - la méthode GetSQL nous donne le SQL du noeud en dessous
    SSQL:='SELECT * FROM YGEDDICO AS YG2 WHERE EXISTS('
         + GetTabletteSQL('YYGEDNIVEAU2', DataTypeLink.GetSql(TGN1.CodeGed, False)+' AND (YG2.YGD_CODEGED=YGEDDICO.YGD_CODEGED)',True )+')'
         + ' AND YGD_ARBOGED="'+SArboGed+'"'
         + ' ORDER BY YGD_TRIGED';

    TobNiv2 := Tob.Create ('Ged Niveau 2', nil, -1);
    TobNiv2.LoadDetailFromSQL (SSQL);

    try
     for j := 0 to TobNiv2.Detail.Count-1 do
      begin
       Niv2Utile := False;
       T         := TobNiv2.Detail [j];

       TGN2         := TGedDPNode.Create;
       TGN2.CodeGed := T.GetValue('YGD_CODEGED');
       TGN2.TypeGed := T.GetValue('YGD_TYPEGED');
       TGN2.SFileGUID := '';
       TGN2.SDocGUID  := '';

       TGN2.TagMenu := T.GetValue('YGD_TAGMENU');
       TGN2.IconId  := cImgClasseurOuvert;
       TGN2.IconSel := cImgClasseurOuvert;

       TGN2.Caption := T.GetValue('YGD_LIBELLEGED');
       TGN2.CDPubli := T.GetString ('YGD_CDPUBLI');

       TN2               := TvPublication.Items.AddChildObject(TN1, TGN2.Caption, TGN2);
       TN2.ImageIndex    := TGN2.IconId;
       TN2.SelectedIndex := TGN2.IconSel;

       //--- Niveau 3 - la méthode GetSQL nous donne le SQL du noeud en dessous
       SSQL:='SELECT * FROM YGEDDICO AS YG3 WHERE EXISTS('
            + GetTabletteSQL('YYGEDNIVEAU3', DataTypeLink2.GetSql(TGN2.CodeGed, False)
            + ' AND (YG3.YGD_CODEGED=YGEDDICO.YGD_CODEGED)',True )+')'
            + ' AND YGD_ARBOGED="'+SArboGed+'"'
            + ' ORDER BY YGD_TRIGED';
       TobNiv3 := Tob.Create('Ged Niveau 3', nil, -1);
       TobNiv3.LoadDetailFromSQL(SSQL);

       try
        for k:=0 to TobNiv3.Detail.Count-1 do
         begin
          Niv3Utile := False;
          T    := TobNiv3.Detail[k];
          TGN3 := TGedDPNode.Create;
          TGN3.CodeGed := T.GetValue('YGD_CODEGED');
          TGN3.TypeGed := T.GetValue('YGD_TYPEGED');
          TGN3.SFileGUID:='';
          TGN3.SDocGUID:='';
          TGN3.TagMenu := T.GetValue('YGD_TAGMENU');
          TGN3.IconId := cImgIntercalOuvert;
          TGN3.IconSel := cImgIntercalOuvert;
          TGN3.Caption := T.GetValue('YGD_LIBELLEGED');
          TGN3.CDPubli := T.GetString ('YGD_CDPUBLI');

          TN3               := TvPublication.Items.AddChildObject(TN2, TGN3.Caption, TGN3);
          TN3.ImageIndex := TGN3.IconId;
          TN3.SelectedIndex := TGN3.IconSel;

          if AjouteNoeudDocuments (TvPublication,TN3) then Niv3Utile := TRUE;
          if Niv3Utile = FALSE then
           begin
            TGedDPNode (TN3.Data).Free;
            TN3.Delete;
           end
          else
           Niv2Utile := TRUE;
         end;
       finally
        TobNiv3.Free;
       end;

      if AjouteNoeudDocuments (TvPublication,TN2) then Niv2Utile := TRUE;
      if Niv2Utile = FALSE then
        begin
         TGedDPNode (TN2.Data).Free;
         TN2.Delete;
        end
       else
        Niv1Utile := TRUE;
      end;
    finally
     TobNiv2.Free;
    end;

    if AjouteNoeudDocuments (TvPublication,TN1) then
     Niv1Utile := TRUE;
     if Niv1Utile = FALSE then
      begin
       TGedDPNode (TN1.Data).Free;
       TN1.Delete;
      end;
   end;
 finally
  TobNiv1.Free;
 end;

 AjouteNoeudDocuments (TvPublication,nil);
 TvPublication.FullExpand;
 TvPublication.Items.EndUpdate;
end;

procedure TFPublicationCD.LoadJaquetteList;
var
   TOBJaq    :TOB;
   i         :integer;
begin
     // Liste des états de jaquette (nature DJA)
     CBJaquette.Clear;
     CBJaquette.Items.Add ('<aucune>');
     TOBJaq := TOB.Create ('les jaquettes', nil, -1);
     try
        TOBJaq.LoadDetailFromSQL ('SELECT MO_CODE,MO_LIBELLE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="DJA" ORDER BY MO_CODE,MO_LIBELLE');
        for i := 0 to TOBJaq.Detail.Count - 1 do
            CBJaquette.Items.Add (TOBJaq.Detail [i].GetString ('MO_CODE') + ' - ' + TOBJaq.Detail [i].GetString ('MO_LIBELLE'));
     finally
            if CBJaquette.Items.Count > 1 then
               CBJaquette.ItemIndex := 1;
            TOBJaq.Free;
     end;
end;

//------------------------------------
//--- Nom : TVPublication_Expanding
//------------------------------------
procedure TFPublicationCD.TVPublication_Expanding (Sender:TObject; Node:TTreeNode; var AllowExpansion:Boolean);
var TN            : TTreeNode;
    TGNNode, TGN  : TGedDPNode;
    TobDoc, T     : TOB;
    i             : Integer;
    TvPublication : TTreeView;
begin
 if Node.Data = Nil then exit;

 if (TControl(Sender).Name='TVPublication_DP') then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 TGNNode := TGedDPNode (Node.Data);

 //--- Racine d'un ensemble de documents
 if ((TGNNode.TypeGed='DCS') {or (TGNNode.TypeGed='CLS')} or (TGNNode.TypeGed='ITL')) and (Node.HasChildren) then
  begin
   //--- Vire 1er noeud factice (servait juste à avoir le +) et ajoute les documents
   TN := Node.GetFirstChild;
   if TN.Text = 'Dummy' then
    begin
     TN.Delete;
     //--- Ajoute les noeuds documents
     TobDoc := TOB.Create ('Les documents', nil, -1);

     //--- INNER car il y a toujours un document, même s'il n'a pas de fichiers
     try
      TobDoc.LoadDetailFromSQL ( 'SELECT DPDOCUMENT.*, YDOCUMENTS.*, YDOCFILES.*, YFILES.YFI_FILENAME '
                               + 'FROM DPDOCUMENT INNER JOIN YDOCUMENTS ON DPD_DOCGUID=YDO_DOCGUID '
                               + 'LEFT JOIN YDOCFILES ON YDO_DOCGUID=YDF_DOCGUID '
                               + 'LEFT JOIN YFILES ON YDF_FILEGUID=YFI_FILEGUID '
                               + 'WHERE DPD_NODOSSIER="'+VH_Doss.NoDossier+'" '
                               + 'AND DPD_CODEGED="'+TGNNode.CodeGed+'" '
                               + 'AND (YDF_FILEROLE IS NULL OR YDF_FILEROLE="PAL") '      // uniquement le fichier Principal
                               + 'ORDER BY YDO_DATECREATION DESC'); // #### Quel classement ?

      //--- Voir pour niveau intermédiaire de regroupt selon un chp quelconque, par exemple YDO_ANNEE (=millésime)
      for i := 0 to TobDoc.Detail.Count-1 do
       begin
        T             := TobDoc.Detail [i];
        TGN           := TGedDPNode.Create;
        TGN.CodeGed   := TGNNode.CodeGed;
        TGN.SFileGUID := '';
        TGN.SDocGUID  := T.GetValue('DPD_DOCGUID');
        TGN.TagMenu   := 0;
        TGN.Url       := '';

        // Il y a bien un fichier associé au document
        if T.GetString('YDF_DOCGUID')<>'' then
         begin
          // donc c'est un type 'DOC' pour le noeud Ged
          // $$$ JP 16/05/06 - ne pas typer le doc si pas droit d'accès à ce doc: FQ 11049
          if (T.GetString ('YDO_UTILISATEUR') <> V_PGI.User) and (JaiLeDroitConceptBureau (ccConsulterDocumentAutre) = FALSE) then
           begin
            TGN.TypeGed := 'DNO';
            TGN.IconId  := cImgCDPubliInactif;
           end
          else
           begin
            TGN.TypeGed := 'DOC';
            TGN.IconId  := cImgCDPubliNot; // Document;
           end;
          TGN.SFileGUID := T.GetValue('YDF_FILEGUID');
         end
        else
         begin
          // $$$ JP 19/12/06: URL ou COMmentaire
          if T.GetString ('YDO_DOCTYPE') = 'URL' then
           begin
            TGN.TypeGed := 'URL';
            TGN.Url     := T.GetString ('YDO_URL');
            TGN.IconId  := cImgInternetExplorer;
           end
          else
           begin
            TGN.TypeGed := 'COM';
            TGN.IconId := cImgCommentaire;
           end;
         end;

        // $$$ JP 17/05/06 TGN.IconSel := TGN.IconId ;
        TGN.Caption := T.GetValue ('YDO_LIBELLEDOC');
        // $$$ JP 16/05/06 - FQ 11050: publication d'un élément si c'est un document accessible
        // $$$ JP 17/05/06 - FQ 10877: filtrage sur abonnement fait dans FiltrerSelection
        //if TGN.TypeGed = 'DOC' then
        //TGN.CDPubli := TGNNode.CDPubli
        //else
        TGN.CDPubli := '';

        TGN.Auteur  := T.GetString ('YDO_AUTEUR');
        TGN.Nature  := T.GetString ('YDO_NATDOC');
        TGN.Mois    := T.GetString ('YDO_MOIS');
        TGN.Annee   := T.GetString ('YDO_ANNEE');

        // $$$ JP 17/05/06 - FQ 10877: filtrage sur abonnement fait dans FiltrerSelection: icon spécifié plus haut
        {if TGN.CDPubli <> '' then
        begin
        TGN.IconId  := cImgCDPubli;
        TGN.IconSel := cImgCDPubli;
        end
        else
        begin
        TGN.IconId  := cImgDocument;
        TGN.IconSel := cImgDocument;
        end;}

        // création du noeud
        TGN.IconSel      := TGN.IconId;
        TN               := TvPublication.Items.AddChildObject (Node, TGN.Caption, TGN);
        TN.ImageIndex    := TGN.IconId;
        TN.SelectedIndex := TGN.IconSel;
       end;
      finally
       TobDoc.Free;
     end;
    end;
  end;
end;

//----------------------------
//--- Nom : BDefaut_DPClick
//----------------------------
procedure TFPublicationCD.BDefaut_DPClick(Sender: TObject);
begin
 if PgiAsk ('Confirmez-vous la reprise des filtres définis à l''étape précédente (pour toute la publication) ?') = mrYes then
  begin
   try
    if (TControl(Sender).name='TVPublication_DP') then
     FilterSelection (TvPublication_DP,'D')
    else
     FilterSelection (TvPublication_DA,'A');
    SourisSablier;
   finally
    SourisNormale;
   end;
  end;
end;

procedure TFPublicationCD.SelectNode (TvPublication : TTreeView; Node:TTreeNode);
var
   TGN        :TGedDPNode;
   strExplic  :string;
   strPubli   :string;
begin
     strExplic := '';
     if TvPublication.Selected <> nil then
     begin
          TGN := TGedDPNode (TvPublication.Selected.Data);
          if (TGN <> nil) then
          begin
               strPubli := TGN.CDPubli;
               strExplic := TGN.Caption + ':' + #10;
               if (TGN.TypeGed <> 'ARM') and (TGN.TypeGed <> 'CLS') and (TGN.TypeGed <> 'ITL') and (TGN.TypeGed <> 'DCS') then
               begin
                    // $$$ JP 17/05/06 FQ 11049: il faut savoir pourquoi un doc' ne réagit aux demandes de publication
                    if TGN.TypeGed = 'DOC' then
                    begin
                         if strPubli <> '' then
                             strExplic := strExplic + 'publié'
                         else
                             strExplic := strExplic + 'non publié';
                    end
                    else if TGN.TypeGed = 'DNO' then
                         strExplic := strExplic + 'non publiable (droits insuffisants)'
                    else if TGN.TypeGed = 'COM' then
                         strExplic := strExplic + 'non publiable (commentaire)'
                    else
                         strExplic := strExplic + 'élément ged inconnu';
               end
               else
               begin
                    if (strPubli <> '') and (strPubli <> 'X') then
                    begin
                         strExplic := strExplic + 'documents groupés par ';
                         if Copy (strPubli, 2, 3) = 'M&A' then
                                 strExplic := strExplic + 'mois/année'
                            else if Copy (strPubli, 2, 3) = 'AUT' then
                                 strExplic := strExplic + 'auteur'
                            else if Copy (strPubli, 2, 3) = 'NAT' then
                                 strExplic := strExplic + 'nature';
                    end
                    else
                        strExplic := strExplic + 'contenu non groupé';
               end;
          end;
     end;
     LExplic_DP.Caption := strExplic;
end;

procedure TFPublicationCD.ChangeCDPubli (TvPublication : TTreeView; TN:TTreeNode; strCDPubli:string=#0; bOnlyDocs:boolean=TRUE; bPropagate:boolean=TRUE);
var
   TGN     :TGedDPNode;
   TNChild :TTreeNode;
begin
     if TN <> nil then
     begin
          TGN := TGedDPNode (TN.Data);
          if TGN <> nil then
          begin
               // Début modification sur l'arborescence
               TvPublication.Items.BeginUpdate;

               // On gère la mise à jour sur les noeuds internes
               if TGN.SDocGUID = '' then
               begin
                    // Pour les types de regroupements, on affecte au noeud interne
                    if bOnlyDocs = FALSE then
                       TGN.CDPubli := strCDPubli;

                    // Propagation si demandé
                    if bPropagate = TRUE then
                    begin
                         TNChild := TN.GetFirstChild;
                         while TNChild <> nil do
                         begin
                              // Changement état publication sur l'enfant
                              ChangeCDPubli (TvPublication, TNChild, strCDPubli, bOnlyDocs);

                              // Enfant suivant
                              TNChild := TNChild.GetNextSibling;
                         end;

                    end;
               end
               else
               begin
                    // $$$ JP 16/05/06 - FQ 11050: changement permit que sur les documents accessibles
                    if TGN.TypeGed = 'DOC' then
                    begin
                         // On détermine le nouvel état de publication (#0 signifie "permutation publié/non publié")
                         if strCDPubli = #0 then
                         begin
                              if TGN.CDPubli = '' then
                                  strCDPubli := 'X'
                              else
                                  strCDPubli := '';
                         end;

                         // On ne transmet pas les codes regroupement aux documents (ils ne concernent que les noeuds internes)
                         if bOnlyDocs = TRUE then
                         begin
                              // On fixe ce nouvel état (avec màj icônes dans l'arborescence)
                              TGN.CDPubli := strCDPubli;
                              if TGN.TypeGed = 'ARM' then
                                   TGN.IconId := cImgArmoireFerme
                              else if TGN.TypeGed = 'CLS' then
                                   TGN.IconId := cImgClasseurFerme
                              else if TGN.TypeGed = 'DCS' then
                                   TGN.IconId := cImgDocuments
                              else if TGN.CDPubli = 'X' then
                                   TGN.IconId := cImgCDPubli
                              else
                                   TGN.IconId := cImgCDPubliNot; // Document;
                              TGN.IconSel      := TGN.IconId;
                              TN.ImageIndex    := TGN.IconId;
                              TN.SelectedIndex := TGN.IconSel;
                         end;
                    end;
               end;

               // Fin modification sur l'arborescence
               TvPublication.Items.EndUpdate;
          end;
     end;
end;

procedure TFPublicationCD.PopStructurePopup (Sender:TObject);
var TGN           : TGedDPNode;
    strPubli      : string;
    TvPublication : TTreeView;
begin
 inherited;

 if (TControl(Sender).name='TVPublication_DP') then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 if TvPublication.Selected <> nil then
  begin
   TGN := TGedDPNode (TvPublication.Selected.Data);
          if TGN <> nil then
          begin
               strPubli := TGN.CDPubli;
               if (TGN.TypeGed = 'ARM') or (TGN.TypeGed = 'CLS') or (TGN.TypeGed = 'DCS') then
               begin
                    MPublier.Caption    := 'TOUT publier';
                    MPasPublier.Caption := 'ne RIEN publier';
                    MRegroupe.Visible   := TRUE;
                    MPublier.Visible    := TRUE;
                    MPasPublier.Visible := TRUE;

                    if Copy (strPubli, 2, 3) = 'M&A' then
                    begin
                         MNoRegroupe.Checked := FALSE;
                         MMoisAnnee.Checked  := TRUE;
                         MNature.Checked     := FALSE;
                         MAuteur.Checked     := FALSE;
                    end
                    else if Copy (strPubli, 2, 3) = 'AUT' then
                    begin
                         MNoRegroupe.Checked := FALSE;
                         MMoisAnnee.Checked  := FALSE;
                         MNature.Checked     := FALSE;
                         MAuteur.Checked     := TRUE;
                    end
                    else if Copy (strPubli, 2, 3) = 'NAT' then
                    begin
                         MNoRegroupe.Checked := FALSE;
                         MMoisAnnee.Checked  := FALSE;
                         MNature.Checked     := TRUE;
                         MAuteur.Checked     := FALSE;
                    end
                    else
                    begin
                         MNoRegroupe.Checked := TRUE;
                         MMoisAnnee.Checked  := FALSE;
                         MNature.Checked     := FALSE;
                         MAuteur.Checked     := FALSE;
                    end;
               end
               else
               begin
                    MRegroupe.Visible := FALSE;
                    if strPubli <> '' then
                    begin
                         MPasPublier.Caption := 'ne PAS publier';
                         MPublier.Visible    := FALSE;
                         MPasPublier.Visible := TRUE;
                    end
                    else
                    begin
                         MPublier.Caption    := 'Publier';
                         MPublier.Visible    := TRUE;
                         MPasPublier.Visible := FALSE;
                    end;
               end;
          end;
     end;
end;

//--------------------------
//--- Nom : MPublierClick
//--------------------------
procedure TFPublicationCD.MPublierClick (Sender:TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TSDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 //--- On demande confirmation pour les noeuds internes
 if TvPublication.Selected.HasChildren = TRUE then
  if PgiAsk ('Confirmez-vous la publication du contenu de ' + TvPublication.Selected.Text + ' ?') <> mrYes then
   exit;

 ChangeCDPubli (TvPublication,TvPublication.Selected, 'X');
 SelectNode (TvPublication,TvPublication.Selected);
end;


//-----------------------------
//--- Nom : MPasPublierClick
//-----------------------------
procedure TFPublicationCD.MPasPublierClick (Sender:TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TsDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 //--- On demande confirmation pour les noeuds internes
 if TvPublication.Selected.HasChildren = TRUE then
  if PgiAsk ('Confirmez-vous la NON publication du contenu de ' + TvPublication.Selected.Text + ' ?') <> mrYes then
   exit;

 ChangeCDPubli (TvPublication,TvPublication.Selected, '');
 SelectNode (TvPublication,TvPublication.Selected);
end;

//---------------------------------
//--- Nom : TvPublication_Change
//---------------------------------
procedure TFPublicationCD.TVPublication_Change (Sender:TObject; Node:TTreeNode);
var TvPublication : TTreeView;
begin
 inherited;
 if (TControl(Sender).name='TVPublication_DP') then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;
 SelectNode (TvPublication, Node);
end;

//---------------------------------
//--- Nom : TvPublication_Change
//---------------------------------
procedure TFPublicationCD.ChangeReg (TvPublication : TTreeView; StrNewReg : string);
var TGN           : TGedDPNode;
    bPropagate    : boolean;
begin
 if TvPublication.Selected <> nil then
  begin
   TGN := TGedDPNode (TvPublication.Selected.Data);
   if TGN <> nil then
    begin
     //--- On demande confirmation de propagation sur une armoire
     bPropagate := FALSE;
     if TGN.TypeGed = 'ARM' then
      if PgiAsk ('Désirez-vous propager la modification aux classeurs de ' + TGN.Caption + ' ?') = mrYes then
       bPropagate := TRUE;

     if TGN.TypeGed = 'CLS' then
      if PgiAsk ('Désirez-vous propager la modification aux intercalaires de ' + TGN.Caption + ' ?') = mrYes then
       bPropagate := TRUE;

     //--- Modification de l'état de regroupement, avec ou sans propagation
     ChangeCDPubli (TvPublication,TvPublication.Selected, strNewReg, FALSE, bPropagate);
     SelectNode (TvPublication,TvPublication.Selected);
    end;
  end;
end;

//-----------------------------
//--- Nom : MNoRegroupeClick
//-----------------------------
procedure TFPublicationCD.MNoRegroupeClick (Sender:TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TsDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 ChangeReg (TvPublication,'X');
end;

//-----------------------------
//--- Nom : MMoisAnneeClick
//-----------------------------
procedure TFPublicationCD.MMoisAnneeClick (Sender:TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TsDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 ChangeReg (TvPublication,'+M&A');
end;

//-----------------------------
//--- Nom : MNatureClick
//-----------------------------
procedure TFPublicationCD.MNatureClick(Sender: TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TsDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 ChangeReg (TvPublication,'+NAT');
end;

//-----------------------------
//--- Nom : MAuteurClick
//-----------------------------
procedure TFPublicationCD.MAuteurClick (Sender:TObject);
var TvPublication : TTreeView;
begin
 if (P.ActivePage=TsDossierPermanent) then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 ChangeReg (TvPublication,'+AUT');
end;

procedure TFPublicationCD.BApercuLogoClick (Sender: TObject);
begin
     inherited;

     if FileExists (ELogoFile.Text) = TRUE then
        ShellExecute (0, pchar ('open'), pchar (ELogoFile.Text), nil, nil, SW_RESTORE);
end;

procedure TFPublicationCD.BSelectLogoClick (Sender:TObject);
var
   LogoFileDlg    :TOpenDialog;
begin
     inherited;

     LogoFileDlg := TOpenDialog.Create (Self);
     try
        LogoFileDlg.Filter     := 'Toutes images (jpeg, bmp, gif)|*.jpg;*.jpeg;*.bmp;*.gif|Images Expert Group (*.jpg;*.jpeg)|*.jpg;*.jpeg|Images Bitmap (*.bmp)|*.bmp|Images Compuserve (*.gif)|*.gif';
        LogoFileDlg.Options    := [ofFileMustExist, ofHideReadOnly, ofNoChangeDir];
        if LogoFileDlg.Execute = TRUE then
           ELogoFile.Text := LogoFileDlg.FileName
     finally
            LogoFileDlg.Free;
     end;
end;

procedure TFPublicationCD.BSelectPathClick(Sender: TObject);
var
  strPath    :string;
begin
     strPath := EPathCible.Text;
     if SelectDirectory (strPath, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) = TRUE then
     begin
          if UpperCase (Copy (strPath, Length (strPath)-8, 9)) <> 'CD-CLIENT' then
             strPath := strPath + '\CD-CLIENT';
          EPathCible.Text := strPath;
     end;
end;

procedure TFPublicationCD.BApercuSiteCabClick(Sender: TObject);
var
   strSiteCab    :string;
begin
     strSiteCab := Trim (ESiteCab.Text);
     if strSiteCab <> '' then
        ShellExecute (0, pchar ('open'), pchar ('http://' + strSiteCab), nil, nil, SW_RESTORE);
end;

procedure TFPublicationCD.bPrecedentClick(Sender: TObject);
begin
{     if (P.ActivePageIndex = 2) and (m_bFiltered = TRUE) then
        if PgiAsk ('Le filtre déjà défini va être annulé'#10' Confirmez-vous le retour sur la page de sélection du filtre ?') = mrYes then
           BuildPublication;}

     inherited;
end;

procedure TFPublicationCD.bSuivantClick (Sender:TObject);
begin
     // Sur dernier onglet, on remplit les paramètres de génération
     if P.ActivePageIndex = 5 then
     begin
          // Si pas de modèle défini, impossible de continuer
          if DirectoryExists (EPathModele.Text) = FALSE then
          begin
               PgiInfo ('Le répertoire modèle est obligatoire: veuillez le sélectionner');
               exit;
          end;

          // Idem pour l'emplacement (+ vérif qu'on est pas sur un disque PGI, ni dans le répetoire PGI00)
          if EPathCible.Text = '' then
          begin
               PgiInfo ('Le répertoire cible de la publication est obligatoire: veuillez le sélectionner');
               exit;
          end;
          if UpperCase (Copy (EPathCible.Text, 2, 7)) = ':\PGI00' then
          begin
               PgiInfo ('Le répertoire cible "' + EPathCible.Text + '" appartient au répertoire système Cegid' + #10 + ' Il n''est pas utilisable pour la publication');
               exit;
          end;

          // Affichage du résumé des options avant traitement
          EResume.Lines.Clear;
          EResume.Lines.Add ('  - Répertoire cible ' + #9#9 + EPathCible.Text);
          EResume.Lines.Add ('  - Répertoire modèle ' + #9 + EPathModele.Text);
          if CBGraveurs.ItemIndex > 0 then
             EResume.Lines.Add ('  - Unité de gravure CD ' + #9 + CBGraveurs.Text);
          if ELogoFile.Text <> '' then
             EResume.Lines.Add ('  - Logo cabinet ' + #9#9 + ExtractFileName (ELogoFile.Text));
          if ETitleHtml.Text <> '' then
             EResume.Lines.Add ('  - Titre publication ' + #9#9 + ETitleHtml.Text);
          if ECommentaire.Text <> '' then
             EResume.Lines.Add ('  - Commentaire ' + #9#9 + ECommentaire.Text);
          EResume.Lines.Add ('');
          EResume.Lines.Add ('Veuillez cliquer sur "Publier" pour générer la publication');

          bFin.Enabled := TRUE;
     end
     else
     begin
          bFin.Enabled := FALSE;

          // Application du filtre si on sort de la page de sélection des filtres
          if P.ActivePageIndex = 1 then
          begin
               if (TControl(Sender).name='TvPublication_DP') then
                begin
                 EMoisAnnee_DP.Text := Trim (EMoisAnnee_DP.Text);
                 EAuteur_DP.Text    := Trim (EAuteur_DP.Text);
                 FilterSelection (TvPublication_DP,'D');
                end
               else
                begin
                 EMoisAnnee_DA.Text := Trim (EMoisAnnee_DA.Text);
                 EAuteur_DA.Text    := Trim (EAuteur_DA.Text);
                 FilterSelection (TvPublication_DA,'A');
                end;
          end;
     end;

     inherited;
end;

function TFPublicationCD.GetResume (iNumLigne:integer):string;
var
   TN      :TTreeNode;
   i       :integer;
begin
     Result := '';

     // On traite l'ensemble des armoires, jusqu'à trouvé celle dont l'indice est précisé
     TN := TvPublication_DP.Items.GetFirstNode;
     i  := 1;
     while (TN <> nil) and (Result = '') do
     begin
          if iNumLigne = i then
              Result := TN.Text
          else
          begin
               TN := TN.GetNextSibling;
               i  := i + 1;
          end;
     end;
end;

procedure TFPublicationCD.LoadData (strLogoFile:string);
var
   TOBDossier    :TOB;
begin
     m_TOBData := TOB.Create ('Donnees CD-CLIENT', nil, -1);

     // Données concernant le cabinet
     m_TOBData.AddChampSupValeur ('CAB_NOM', V_PGI.NomSociete);
     m_TOBData.AddChampSupValeur ('CAB_LOGO', strLogoFile);
     m_TOBData.AddChampSupValeur ('CAB_SITE', Trim ('http://' + ESiteCab.Text));

     // Données concernant le dossier
     TOBDossier := TOB.Create ('le dossier', nil, -1);
     try
        TOBDossier.LoadDetailFromSQL ('SELECT ANN_NOM1,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALCP,ANN_ALVILLE FROM ANNUAIRE,DOSSIER WHERE DOS_NODOSSIER="' + VH_DOSS.NoDossier + '" AND ANN_GUIDPER=DOS_GUIDPER');
        if TOBDossier.Detail.Count > 0 then
        begin
             m_TOBData.AddChampSupValeur ('SOC_NOM',   Trim (TOBDossier.Detail [0].GetString ('ANN_NOM1')));
             m_TOBData.AddChampSupValeur ('SOC_ADR1',  Trim (TOBDossier.Detail [0].GetString ('ANN_ALRUE1')));
             m_TOBData.AddChampSupValeur ('SOC_ADR2',  Trim (TOBDossier.Detail [0].GetString ('ANN_ALRUE2') + ' ' + TOBDossier.Detail [0].GetString ('ANN_ALRUE3')));
             m_TOBData.AddChampSupValeur ('SOC_CP',    Trim (TOBDossier.Detail [0].GetString ('ANN_ALCP')));
             m_TOBData.AddChampSupValeur ('SOC_VILLE', Trim (TOBDossier.Detail [0].GetString ('ANN_ALVILLE')));
        end
        else
        begin
             m_TOBData.AddChampSupValeur ('SOC_NOM',   '???');
             m_TOBData.AddChampSupValeur ('SOC_ADR1',  '???');
             m_TOBData.AddChampSupValeur ('SOC_ADR2',  '???');
             m_TOBData.AddChampSupValeur ('SOC_CP',    '???');
             m_TOBData.AddChampSupValeur ('SOC_VILLE', '???');
        end;
     finally
            TOBDossier.Free;
     end;

     // Données concernant la publication
     m_TOBData.AddChampSupValeur ('PUB_TITRE',       Trim (ETitleHtml.Text));
     m_TOBData.AddChampSupValeur ('PUB_COMMENTAIRE', Trim (ECommentaire.Text));
end;

procedure TFPublicationCD.DeleteDirectory (strPath:string);
var
   FileList    :TStringList;
   DirList     :TStringList;
   sr          :TSearchRec;
   strFile     :string;
   i           :integer;
begin
     FileList := TStringList.Create;
     DirList  := TStringList.Create;
     try
        // 1 - Enumération des éléments à supprimer (pour un rép., on supprime son contenu)
        if FindFirst (strPath + '\*.*', faAnyFile, sr) = 0 then
        begin
             repeat
                   strFile := strPath + '\' + sr.Name;
                   if (sr.Attr and faDirectory) > 0 then
                   begin
                        if (sr.Name <> '.') and (sr.Name <> '..') then
                        begin
                             DeleteDirectory (strFile);
                             DirList.Add (strFile);
                        end;
                   end
                   else
                   begin
                        if ((sr.Attr and faSysFile) = 0) and ((sr.Attr and faVolumeID) = 0) then
                           FileList.Add (strFile);
                   end;
             until FindNext (sr) <> 0;
             FindClose (sr);
        end;

        // 2 - Suppression des éléments stockés (normalement, les rép. son déjà vides)
        for i := 0 to DirList.Count-1 do
            RemoveDir (DirList [i]);
        for i := 0 to FileList.Count-1 do
            DeleteFile (FileList [i]);
     finally
            DirList.Free;
            FileList.Free;
     end;
end;

procedure TFPublicationCD.bFinClick (Sender:TObject);
var
   strLogoFile   :string;
   AutoLines     :TStringList;
begin
     inherited;

     // Si déjà générer, on ferme
     if m_bGenerated = TRUE then
     begin
          ModalResult := mrOK;
          exit;
     end;

     // Demande confirmation avant de faire la génération
     if PgiAsk ('Confirmez-vous la génération de la publication?') <> mrYes then
        exit;

     // Demande confirmation si répertoire cible existe déjà
     if DirectoryExists (EPathCible.Text) = TRUE then
     begin
          if PgiAsk ('Le répertoire "' + EPathCible.Text + '" existe déjà.' + #10 + ' Le contenu de ce répertoire va être totalement supprimé' + #10 + ' Désirez-vous continuer le traitement ?') = mrYes then
          begin
               if PgiAsk ('Tous les fichiers de "' + EPathCible.Text + '" vont être supprimés.' + #10 + ' Désirez-vous annuler le traitement ?') = mrYes then
                   exit
               else
                   DeleteDirectory (EPathCible.Text);
          end
          else
              exit;
     end
     else
     begin
          // Création du répertoire
          if ForceDirectories (EPathCible.Text) = FALSE then
          begin
               PgiInfo ('Impossible de créer le répertoire "' + EPathCible.Text + '"');
               exit;
          end;
     end;

     // Début de traitement
     SourisSablier;

     // Nom du fichier logo
     strLogoFile := EPathCible.Text + '\cd_logo' + ExtractFileExt (ELogoFile.Text);

     // On stocke toutes les données dans une TOB unique (nom dossier, cabinet, logo...)
     LoadData (strLogoFile);

     // Générations fichiers et sommaire html
     try
        // On se sert du résumé pour le log de traitement
        EResume.Clear;
        EResume.Lines.Add ('Début de génération de la publication HTML');
        EResume.Lines.Add ('');
        EResume.Refresh;

        // Copie de tous les fichiers contenu dans le répertoire modèle (y compris sous-répertoires), pour ne pas oublier des fichiers liés à index.html
        EResume.Lines.Add ('  * Copie des fichiers modèles dans le répertoire cible');
        EResume.Refresh;
        CopyDirectory (EPathModele.Text, EPathCible.Text, '*.*', TRUE, FALSE);

        // Copie du fichier logo dans le répertoire de préparation
        if ELogoFile.Text <> '' then
        begin
             EResume.Lines.Add ('  * Copie du fichier logo dans le répertoire cible');
             EResume.Refresh;
             if CopyFile (pchar (ELogoFile.Text), pchar (strLogoFile), FALSE) = FALSE then
                strLogoFile := '';
        end;

        // Génération du fichier html "sommaire", en fonction du modèle et des infos cabinet/dossier
        StrPathCible:= EPathCible.Text;
        CreateHtml;
        EResume.Refresh;

        // Création du fichier autorun
        EResume.Lines.Add ('  * Création du fichier autorun.inf (lancement automatique du CD-ROM)');
        AutoLines := TStringList.Create;
        try
           AutoLines.Add ('[Autorun]');
           AutoLines.Add ('shellexecute=index.html');

           AutoLines.SaveToFile (EPathCible.Text + '\autorun.inf');
        finally
               AutoLines.Free;
        end;

        // On a terminé la génération, mais on ne quitte pas l'assistant
        EResume.Lines.Add ('');
        EResume.Lines.Add ('Fin de la génération de la publication HTML');
        m_bGenerated         := TRUE;
        BApercuPubli.Enabled := TRUE;
        bPrecedent.Enabled   := FALSE;
        bFin.Caption         := 'Fermer';
        if CBGraveurs.ItemIndex > 0 then
        begin
             BGraveCD.Enabled := TRUE;
             EResume.Lines.Add ('');
             EResume.Lines.Add ('Pour graver la publication:');
             EResume.Lines.Add ('     - vérifier que l''option d''enregistrement de l''unité de gravure est activée');
             EResume.Lines.Add ('     - insérer un CD-ROM dans l''unité de gravure');
             EResume.Lines.Add ('     - attendre le chargement total du CD-ROM');
             EResume.Lines.Add ('     - cliquer sur le bouton "Graver"');
        end;
        if CBJaquette.ItemIndex > 0 then
           BImpJaquette.Enabled := TRUE;
     finally
            EResume.Refresh;
            SourisNormale;
     end;
end;

procedure TFPublicationCD.BGraveCDClick (Sender:TObject);
begin
     // Si déjà généré correctement, on lance la gravure ou on quitte (selon si gravure demandée ou pas)
     if m_bGenerated = TRUE then
     begin
          if CBGraveurs.ItemIndex > 0 then
          begin
               // Demande de mettre un CD dans l'unité de gravure
               if PgiAsk ('Veuillez insérer un CD inscriptible dans l''unité ' +  m_BurnCD.DataRecorders.Items [0].DosDevice + ' ' + m_BurnCD.ActiveDataRecorder.Recorder.DisplayName + #10 + '  Veuillez cliquer sur OUI lorsque le CD est est totalement chargé') <> mrYes then
                  exit;

               // Début gravage
               SourisSablier;

               try
                  // Lancement de la gravure sur l'unité sélectionnée
                  m_BurnCD.ActiveDataRecorder.Recorder := TDataRecorder (CBGraveurs.Items.Objects [CBGraveurs.ItemIndex]);
                  EResume.Clear;
                  EResume.Lines.Add ('Début de la gravure CD sur ' + m_BurnCD.DataRecorders.Items [0].DosDevice + ' ' + m_BurnCD.ActiveDataRecorder.Recorder.DisplayName);
                  EResume.Refresh;
                  if BurnCD (EPathCible.Text) = TRUE then
                  begin
                       EResume.Lines.Add ('Fin de la gravure CD');
                       bFin.Caption := 'Quitter';
                       CBGraveurs.ItemIndex := -1;
                  end;
               finally
                      EResume.Refresh;
                      SourisNormale;
               end;
               exit;
          end;
     end;
end;

procedure TFPublicationCD.BImpJaquetteClick (Sender:TObject);
begin
     if CBJaquette.ItemIndex > 0 then
        if PgiAsk ('Confirmez-vous l''impression de la jaquette "' + CBJaquette.Text + '" ?') = mrYes then
           LanceEtat ('E', 'DJA', Copy (CBJaquette.Text, 1, 3), TRUE, FALSE, FALSE, nil, '', 'Jaquette CD', FALSE);
end;

procedure TFPublicationCD.DeleteNoPublished (TvPublication : TTreeView; TN:TTreeNode);
var
   TNChild      :TTreeNode;
   TNNextChild  :TTreeNode;
   TGN          :TGedDPNode;
begin
     // si noeud non défini, on prend le premier de l'arborescence, sinon le premier noeud enfant
     if TN = nil then
         TNChild := TvPublication.Items.GetFirstNode
     else
         TNChild := TN.GetFirstChild;

     // On traite l'ensemble des enfants du noeud
     while TNChild <> nil do
     begin
          // On prend déjà le noeud frère, au cas où celui en cours est supprimé
          TNNextChild := TNChild.GetNextSibling;

          // Noeud GED
          TGN := TGedDPNode (TNChild.Data);

          // Suppression du contenu non publié du noeud en cours, ou du document non publié si c'est un document
          if TGN.SDocGUID = '' then
              DeleteNoPublished (TvPublication, TNChild)
          else
              if TGN.CDPubli = '' then
                 TNChild.Delete;

          // Le noeud frère suivant
          TNChild := TNNextChild;
     end;

     // si le noeud à traiter n'a plus d'enfant
     if (TN <> nil) and (TN.HasChildren = FALSE) then
        TN.Delete;
end;


//----------------------------
//--- Nom : FilterSelection
//----------------------------
procedure TFPublicationCD.FilterSelection (TvPublication : TTreeView; SArboGed : String);
var TNChild          : TTreeNode;
    TGN              : TGedDPNode;
    strParentCDPubli : string;
begin
 //m_bFiltered := FALSE;
 TNChild := TvPublication.Items.GetFirstNode;

 // On traite l'ensemble des enfants du noeud
 while TNChild <> nil do
  begin
   // Noeud GED
   TGN := TGedDPNode (TNChild.Data);

   // Si correspond pas aux critères, on le déselectionne
   if TGN.SDocGUID <> '' then
    begin
     // $$$ JP 17/05/06 FQ 10877: pouvoir filtrer sur tous les documents, et pas que sur les documents
     strParentCDPubli := TGedDPNode (TNChild.Parent.Data).CDPubli;
     if (SArboGed='D') then
      begin
       if (MVBNature_DP.Text <> '') and (Pos (TGN.Nature, MVBNature_DP.Text) = 0) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if (EMoisAnnee_DP.Text <> '') and (TGN.Mois+TGN.Annee <> EMoisAnnee_DP.Text) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if (EAuteur_DP.Text <> '') and (TGN.Auteur <> EAuteur_DP.Text) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if CBAboPub_DP.Checked = TRUE then
        ChangeCDPubli (TvPublication,TNChild, strParentCDPubli)
       else
        ChangeCDPubli (TvPublication,TNChild, 'X');
      end
     else
      begin
       if (MVBNature_DA.Text <> '') and (Pos (TGN.Nature, MVBNature_DA.Text) = 0) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if (EMoisAnnee_DA.Text <> '') and (TGN.Mois+TGN.Annee <> EMoisAnnee_DA.Text) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if (EAuteur_DA.Text <> '') and (TGN.Auteur <> EAuteur_DA.Text) then
        ChangeCDPubli (TvPublication,TNChild, '')
       else if CBAboPub_DA.Checked = TRUE then
        ChangeCDPubli (TvPublication,TNChild, strParentCDPubli)
       else
        ChangeCDPubli (TvPublication,TNChild, 'X');
      end;

    end;

   // Le noeud suivant
   TNChild := TNChild.GetNext;
  end;
end;

function RegroupeDocCompare (Item1:pointer; Item2:pointer):integer;
var
   TN1, TN2                 :TTreeNode;
   TGN1, TGN2               :TGedDPNode;
   strCodeReg1, strCodeReg2 :string;
begin
     // Par défaut égaux (ou bien incomparable)
     Result := 0;

     // Noeuds Treeview à comparer (doivent être issus normalement de même parent)
     TN1 := TTreeNode (Item1);
     TN2 := TTreeNode (Item2);
     if TN1.Parent <> TN2.Parent then
        exit;

     // Noeud GED servant pour la comparaison d'une de leur caractéristique
     TGN1 := TGedDPNode (TN1.Data);
     TGN2 := TGedDPNode (TN2.Data);

     // On compare les champs en fonction du type de regroupement
     strCodeReg1 := TGN1.CodeReg (TGedDPNode (TN1.Parent.Data).CDPubli);
     strCodeReg2 := TGN2.CodeReg (TGedDPNode (TN2.Parent.Data).CDPubli);
     if strCodeReg1 > strCodeReg2 then
          Result := 1
     else if strcodeReg1 < strCodeReg2 then
          Result := -1;
end;

procedure TFPublicationCD.WriteDocuments (HtmlLines:TStringList; DocList:TObjectList);
var
   i                :integer;
   TGN              :TGedDPNode;
   strTitleDoc      :string;
   strDocFileSource :string;
   strDocFileCible  :string;
   strOldRegroup    :string;
   strCurRegroup    :string;
begin
     // On tri la liste des documents en attente selon le critère de regroupement
     DocList.Sort (RegroupeDocCompare);

     // On copie et insère les documents dans cet ordre, en générant un élément à chaque rupture de valeur (regroupement)
     strOldRegroup := '';
     for i := 0 to DocList.Count-1 do
     begin
          // Noeud GED du document à insérer
          TGN := TGedDPNode (TTreeNode (DocList [i]).Data);

          // Si rupture sur valeur de regroupement, on génère un libellé html de regroupement
          strCurRegroup := TGN.DisplayReg (TGedDPNode (TTreeNode (DocList [i]).Parent.Data).CDPubli);
          if strCurRegroup <> strOldRegroup then
          begin
               if strOldRegroup <> '' then
                  HtmlLines.Add ('<br></blockquote>');
               HtmlLines.Add (strCurRegroup + '<blockquote>');

               // Mémorisation valeur de regroupement pour test rupture sur prochain document
               strOldRegroup := strCurRegroup;
          end;

          // Extraction du document
          strDocFileSource := ExtraitDocument (TGN.SDocGUID, strTitleDoc);
          if strDocFileSource <> '' then
          begin
               // Copie du fichier extrait dans le répertoire cible
               strDocFileCible := 'cd_doc' + TGN.SDocGUID + ExtractFileExt (strDocFileSource);
               CopyFile (pchar (strDocFileSource), pchar (strPathCible + '\' + strDocFileCible), FALSE);

               // Référencement dans le sommaire html
               HtmlLines.Add ('<a href="' + strDocFileCible + '">' + TGN.Caption + '</a><br>');
          end
          else
              HtmlLines.Add ('(' + TGN.Caption + ' non disponible)<br>');
     end;

     // Fin de liste: saut de ligne supplémentaire
     HtmlLines.Add ('<br>');

     // On termine le regroupement: retour de tabulation
     if strCurRegroup <> '' then
        HtmlLines.Add ('</blockquote>');

     // On vide la liste des documents qui ne sont plus en attente (ils viennent d'être traités)
     DocList.Clear;
end;

procedure TFPublicationCD.CreateHtml;
var
   HtmlLines    : TStringList;
   HtmlNewLines : TStringList;
   DocList      : TObjectList;
   TN           : TTreeNode;
   i            : integer;
   iPosListe    : integer;
   strLine      : string;
   bResume      : boolean;
   iCurNiveau   : integer;

begin
     EResume.Lines.Add ('  * Génération du sommaire HTML');

     HtmlLines    := TStringList.Create;
     HtmlNewLines := TStringList.Create;
     DocList      := TObjectList.Create (FALSE);
     try

        // Pré-traitement sur la publication: on vire les éléments non publiés (y compris armoires et classeurs)
        DeleteNoPublished (Tvpublication_DP,nil);
        DeleteNoPublished (Tvpublication_DA,nil);

        // $$$ JP 03/10/06: sur .hmod et non plus .html, pb aléatoire d'accès concurrent en sauvegarde (cf SL)
        DeleteFile (strPathCible + '\index.hmod');
        if RenameFile (strPathCible + '\index.html', strPathCible + '\index.hmod') = FALSE then
        begin
             PgiInfo ('impossible de renommer l''index sur ' + strPathCible);
             EResume.Lines.Add ('  * Renommage index.html en index.hmod impossible sur ' + strPathCible);
             exit;
        end;

        // On lit et ré-écrit le fichier html sommaire du répertoire de génération, en changeant simplement les mots clés @@
        HtmlLines.LoadFromFile (strPathCible + '\index.hmod'); // $$$ JP 03/10/06: sur .hmod et non plus .html, pb aléatoire d'accès concurrent en sauvegarde (cf SL)
        for i := 0 to HtmlLines.Count-1 do
        begin
             strLine    := HtmlLines [i];
             iCurNiveau := -1;

             // On construit la liste là où le mot clé @@PUB_LISTE apparait, ou la liste compact si @@PUB_RESUME
             bResume := FALSE;
             iPosListe := Pos ('@@PUB_LISTE', strLine);
             if iPosListe = 0 then
             begin
                  iPosListe := Pos ('@@PUB_RESUME', strLine);
                  bResume := TRUE;
             end;
             if iPosListe > 0 then
             begin
                  // On supprime le mot clé et le suite de la ligne est décalée sur une nouvelle ligne suivante
                  HtmlNewLines.Add (Copy (strLine, 1, iPosListe-1));
                  if bResume = FALSE then
                      Delete (strLine, 1, iPosListe+10)
                  else
                      Delete (strLine, 1, iPosListe+11);


                  //--- On insère le code html décrivant les armoires/classeurs/documents publiés du Dossier permanent
                  TN := TvPublication_DP.Items.GetFirstNode;
                  if (TN<>Nil) then
                   begin
                    if bResume = FALSE then HtmlNewLines.Add ('<H3><Strong>DOSSIER PERMANENT</Strong></H3>');
                    if bResume = TRUE then HtmlNewLines.Add ('<H6><Strong>DOSSIER PERMANENT</Strong></H6>');

                    ConstruireHTMLPublication (TN, HTMLNewLines,DocList,bResume);
                   end;

                  //--- On insère le code html décrivant les armoires/classeurs/documents publiés du Dossier annuel
                  TN := TvPublication_DA.Items.GetFirstNode;
                  if (TN<>nil) then
                   begin
                    if bResume = FALSE then HtmlNewLines.Add ('<H3><Strong>DOSSIER ANNUEL</Strong></H3>');
                    if bResume = TRUE then HtmlNewLines.Add ('<H6><Strong>DOSSIER ANNUEL</Strong></H6>');
                    ConstruireHTMLPublication (TN, HTMLNewLines,DocList,bResume);
                   end;

                  // On termine par la fin de la ligne lue
                  HtmlNewLines.Add (strLine);
             end
             else
             begin
                  // On remplace les mots clés potentiels par les données
                  strLine := StringReplace (strLine, '@@CAB_NOM',         m_TOBData.GetString ('CAB_NOM'),  [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@CAB_LOGO',        ExtractFileName (m_TOBData.GetString ('CAB_LOGO')), [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@CAB_SITE',        m_TOBData.GetString ('CAB_SITE'), [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@SOC_NOM',         m_TOBData.GetString ('SOC_NOM'),   [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@SOC_ADR1',        m_TOBData.GetString ('SOC_ADR1'),  [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@SOC_ADR2',        m_TOBData.GetString ('SOC_ADR2'),  [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@SOC_CP',          m_TOBData.GetString ('SOC_CP'),    [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@SOC_VILLE',       m_TOBData.GetString ('SOC_VILLE'), [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@PUB_TITRE',       '<a name="pub_titre">' + m_TOBData.GetString ('PUB_TITRE') + '</a>', [rfReplaceAll, rfIgnoreCase]);
                  strLine := StringReplace (strLine, '@@PUB_COMMENTAIRE', m_TOBData.GetString ('PUB_COMMENTAIRE'), [rfReplaceAll, rfIgnoreCase]);

                  HtmlNewLines.Add (strLine);
             end;
        end;

        // $$$ JP 03/10/06: on vire de suite le htmllines, au cas où le LoadFromFile aurait encore le handle sur le fichier (cf SL pb client: accès concurrent sur index.html)
        FreeAndNil (HtmlLines);

        // Ecriture dans le répertoire de préparation
        HtmlNewLines.SaveToFile (strPathCible + '\index.html');
     finally
            DocList.Free;
            HtmlNewLines.Free;
            HtmlLines.Free;
     end;
end;

//--------------------------------------
//--- Nom : ConstruireHTMLPublication
//--------------------------------------
procedure TFPublicationCD.ConstruireHTMLPublication (TN : TTreeNode; HTMLNewLines : TStringList; DocList : TObjectList; bResume : Boolean);
var iOldNiveau   : Integer;
    iCurNiveau   : Integer;
    TGN          : TGedDPNode;
    k            : integer;
begin
 iOldNiveau := -1;
 iCurNiveau := -1;

 //--- Pour une jaquette, un seul décalage pour l'ensemble de la liste
 if bResume = TRUE then HtmlNewLines.Add ('<blockquote>');

 while TN <> nil do
  begin
   //--- Noeud GED
   TGN := TGedDPNode (TN.Data);

   if (TGN <> nil) then
    begin
     //--- Identification du niveau hiérarchique de l'élément en cours de traitement
     iCurNiveau := TN.Level;

     //--- S'il y a un changement de niveau: génération docs en attente + màj de la tabulation html
     if (iCurNiveau <> iOldNiveau) and (bResume = FALSE) then
      begin
       //--- S'il y a des documents en attente, on les génère avec le regroupement adéquat (sauf jaquette)
       if DocList.Count > 0 then WriteDocuments (HtmlNewLines, DocList);

       //--- On ajoute ou retranche autant de "tabulation" qu'il est nécessaire
       if (iCurNiveau > iOldNiveau) then
        begin
         for k := iOldNiveau+1 to iCurNiveau do
          HtmlNewLines.Add ('<blockquote>');
        end
       else
        begin
         for k := iOldNiveau-1 downto iCurNiveau do
          HtmlNewLines.Add ('</blockquote>');
        end
      end;

     //--- Mémorisation du niveau en cours pour prochain élément à traiter
     iOldNiveau := iCurNiveau;

     //--- Pour un document: on le stocke dans la liste, afin de trier selon le regroupement choisi
     //--- On affiche un lien hypertexte (lien sur le document, extrait à la volée)
     if (TGN.SDocGUID <> '') and (TGN.CDPubli <> '') and (bResume = FALSE) then
      DocList.Add (TN)
     else
      begin
       //--- Pour une armoire, on met le titre en gras
       if iCurNiveau = 0 then
        if bResume = FALSE then
         HtmlNewLines.Add ('<li><a name="' + TGN.CodeGed + '" style="border-bottom-style:solid; border-bottom-width:1"><b>' + TGN.Caption + '</b></a></li>') //&nbsp;<a href="#pub_titre">(haut)</a></li>')
        else
         HtmlNewLines.Add ('<a href="#' + TGN.CodeGed + '"><b>' + TGN.Caption + '<br></b></a>')
        else
         if bResume = FALSE then HtmlNewLines.Add (TGN.Caption);
      end;
    end;

   //--- Noeud suivant
   TN := TN.GetNext;
  end;

  //--- Il faut termine les documents éventuels en attente
  if bResume = FALSE then
   begin
    WriteDocuments (HtmlNewLines, DocList);

    // Il faut également terminer les tabulations en cours
    for k := iCurNiveau downto 0 do
     HtmlNewLines.Add ('</blockquote>');
   end
  else
   HtmlNewLines.Add ('</blockquote>');

end;

function TFPublicationCD.BurnCD (strPath: string):boolean;
begin
     // Au cas où, mais normalement doit être assigné
     Result := FALSE;
     if m_BurnCD = nil then
        exit;

     if m_BurnCD.ActiveDataRecorder.MediaInfo.isWritable = FALSE then
     begin
          EResume.Lines.Add ('    le CD n''est pas inscriptible ou n''est pas totalement chargé');
          EResume.Lines.Add ('Gravure sur CD annulée');
          exit;
     end;

     // Ajout des éléments et gravure
     try
        m_BurnCD.ClearFiles;
        m_BurnCD.AddFolder (strPath, '');
        m_BurnCD.RecordDisc (FALSE, TRUE);
        Result := TRUE;
     except
           on E:Exception do
           begin
                EResume.Lines.Add ('    erreur: ' + E.Message);
                EResume.Lines.Add ('Gravure sur CD annulée');
           end;
     end;
end;

procedure TFPublicationCD.BSelectModeleClick(Sender: TObject);
var
  strPath    :string;
begin
     inherited;

     // Répertoire modèle par défaut à sélectionner
     strPath := EPathModele.Text;
     if DirectoryExists (strPath) = FALSE then
        strPath := '';

     // Sélection du répertoire modèle
     if SelectDirectory (strPath, [], 0) = TRUE then
     begin
          if (FileExists (strPath + '\index.html') = TRUE) or (PgiAsk ('Le répertoire "' + strPath + '" ne semble pas contenir de modèle html' + #10 + '  Confirmez-vous la sélection de ce répertoire?') = mrYes) then
             EPathModele.Text := strPath;
     end;
end;

procedure TFPublicationCD.BApercuHtmlClick(Sender: TObject);
begin
     if FileExists (EPathModele.Text + '\index.html') = TRUE then
        ShellExecute (0, pchar ('open'), pchar (EPathModele.Text + '\index.html'), nil, nil, SW_RESTORE);
end;


//-----------------------------------
//--- Nom : TvPublication_DblClick
//-----------------------------------
procedure TFPublicationCD.TVPublication_DblClick (Sender:TObject);
var TGN           : TGedDPNode;
    ParGed        : TParamGedDoc;
    TvPublication : TTreeView;
    TOBDoc        : TOB;
begin
 inherited;

 if (TControl(Sender).name='TVPublication_DP') then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 if TvPublication.Selected <> nil then
  begin
   TGN := TGedDPNode (TvPublication.Selected.Data);

   if (TGN.SDocGUID <>'') and (TGN.TypeGed = 'DOC') then // $$$ JP 16/05/06 - FQ 11049
    begin
     ParGed.SDocGUID  := TGN.SDocGUID;
     ParGed.NoDossier := VH_Doss.NoDossier;
     ParGed.CodeGed   := TGN.CodeGed;
     ParGed.SFileGUID := TGN.SFileGUID;
     ParGed.TypeGed   := TGN.TypeGed;

     if ShowNewDocument (ParGed, FALSE, FALSE, TRUE) <> '##NONE##' then
      begin
       // Relecture des 4 propriété principale du document modifié
       TobDoc := TOB.Create ('Le document', nil, -1);
       try
        TobDoc.LoadDetailFromSQL ('SELECT YDO_LIBELLEDOC,YDO_AUTEUR,YDO_NATDOC,YDO_MOIS,YDO_ANNEE FROM YDOCUMENTS WHERE YDO_DOCGUID="' + TGN.SDocGUID + '"');
        if TobDoc.Detail.Count > 0 then
         begin
          TGN.Caption := TobDoc.Detail [0].GetString ('YDO_LIBELLEDOC');
          TGN.Auteur  := TobDoc.Detail [0].GetString ('YDO_AUTEUR');
          TGN.Nature  := TobDoc.Detail [0].GetString ('YDO_NATDOC');
          TGN.Mois    := TobDoc.Detail [0].GetString ('YDO_MOIS');
          TGN.Annee   := TobDoc.Detail [0].GetString ('YDO_ANNEE');

          // Màj du libellé de l'élément en cours (peut avoir changé)
          TvPublication.Selected.Text := TGN.Caption;
         end;
       finally
        TobDoc.Free;
       end;
      end;
    end;
  end;
end;

//-----------------------------------
//--- Nom : TvPublication_DPKeyDown
//-----------------------------------
procedure TFPublicationCD.TVPublication_KeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
var TvPublication : TTreeView;
begin
 inherited;

 if (TControl(Sender).name='TVPublication_DP') then
  TvPublication:=TvPublication_DP
 else
  TvPublication:=TvPublication_DA;

 // Sur barre d'espace, même comportement que marquage à publier/ne pas publier (seulement sur document)
 if Key = VK_SPACE then
  begin
   // On demande confirmation pour les noeuds internes
   if TvPublication.Selected.HasChildren = TRUE then
    if PgiAsk ('Confirmez-vous l''inversion de l''état de publication des documents de ' + TvPublication.Selected.Text + ' ?') <> mrYes then
     exit;

    ChangeCDPubli (TvPublication,TvPublication.Selected);
    SelectNode (TvPublication,TvPublication.Selected);
  end;
end;

procedure TFPublicationCD.BApercuPubliClick (Sender:TObject);
begin
     if FileExists (EPathCible.Text + '\index.html') = TRUE then
        ShellExecute (0, pchar ('open'), pchar (EPathCible.Text + '\index.html'), nil, nil, SW_RESTORE);
end;

procedure TFPublicationCD.BapercuJaquetteClick (Sender:TObject);
begin
     if CBJaquette.ItemIndex > 0 then
        LanceEtat ('E', 'DJA', Copy (CBJaquette.Text, 1, 3), TRUE, FALSE, FALSE, nil, '', 'Modèle de jaquette CD', FALSE);
end;

procedure TFPublicationCD.CBGraveursChange (Sender:TObject);
begin
     CBJaquette.Enabled := CBGraveurs.ItemIndex > 0;
end;

initialization
              FPublicationCD := nil;


end.
