{***********UNITE*************************************************
Auteur  ...... : G.Harlez
Créé le ...... : //2007
Modifié le ... : //2007
Description .. : Source TOF de la FICHE : GRPTRAVAIL ()
Mots clefs ... : TOF;GRPTRAVAIL
//////////////////////////////////////////////////////////////////
Remarques :
  pour la gestion des noeuds dans l'arborescence, j'utilise la propriété
  StateIndex (indice d'état) avec pour valeur :
  -1 : représente un Groupe de Travail
  0  : représente une affectation XXX (en fonction des paramètres renseignés dans FILTREDONNEES)
  1  : représente une affectation Utilisateur
*****************************************************************}
Unit GRPTRAVAIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Menus,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     maineagl,
     eMul,
{$ENDIF}
     HDB,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     DateUtils,
     ExtCtrls,
     HTB97,
     Windows,
     Hpanel,
     Graphics,
     Messages;

// Gestion les icones pour la Treeview (visu icone dans décla)
Const Ico_Grp      = 26-1;
      Ico_GrDenied = 14-1;
      Ico_Delete   = 2-1;
      Ico_Dos      = 23-1;
      Ico_Usr      = 96-1;
      Ico_PtVert   = 47-1;
      Ico_PtRouge  = 46-1;
      Ico_GrpMulti = 66-1;

      Cod_Utilisat = 'YYUTILISAT_SEL';  // Code Fiche Multicritère pour la gestion des users.
      Nat_Utilisat = 'YY';              // Nature Fiche Multicritère pour la gestion des users.

Type
  TTreeItemCode = record
     Code : String ;
     TreeItem : TTreeNode ;
  end;

Type
  TListSearch = record
    StrItemList : String;
    IdxColor : integer;
  end;

Type
  TOrdreTri = ( Croissant,Decroissant );
  TOrdreTris = set of TOrdreTri ;

Type
  TTreeItemCodes = array of TTreeItemCode ;

Type
  TSearchData = array of TListSearch;

Type
  TArrayOfTreeNode = array of TTreeNode;

Type
  TOF_GRPTRAVAIL = Class (TOF)
   {procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnDisplay                ; override ;
    procedure OnCancel                 ; override ;
    procedure OnLoad                   ; override ;}

    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    Obj_TreeVw   : TTreeView;         // Vue hierarchique
    Obj_BtIns    : TToolbarButton97;  // Création d'un groupe de travail
    Obj_BtNIns   : TToolbarButton97;  // Création d'un sous-groupe de travail
    Obj_BtDel    : TToolbarButton97;  // Suppression d'un Item
    Obj_BtRen    : TToolbarButton97;  // Renomme le libelle du groupe
    Obj_BtSearch : TToolbarButton97;  // Recherche les occurences d'un item.
    Obj_BtDnD    : TToolbarButton97;  // Drag n Drop .
    Obj_BtAllDnD : TToolbarButton97;  // Drag n Drop (à ts les groupes).
    Obj_BtSearch_aff : TToolbarButton97; // Recherche un élement de la liste dans la vue hiérarchique.
    Obj_redim    : TSplitter;
    Obj_TxtLib   : THEdit;            // Saisie texte libre. Prompt du radio bouton Obj_RBtn_Dos

    Obj_Pnlst    : THPanel;           // Panel d'affichage des fiches Multicritères.
    Obj_pnldnd   : THPanel;           // Panel contenant le btn d'action DragnDrop + recherche d'un enreg de la liste dans la TV.

    Obj_RBtn_Dos : TRadioButton;      // Radio bouton en fonction de la table de lien (FILTREDONNEES)
    Obj_RBtn_Usr : TRadioButton;      // Radio bouton UTILISATEUR
    Obj_RBtn_All : TRadioButton;      // Radio bouton Tous
    Obj_RBtn_None: TRadioButton;      // Radio bouton Aucun

    WinFindResult : TToolWindow97;
    LstView       : TListView;

    S_Tablien,s_Search,S_TabPrfx,s_Tabliee,
    s_CleTabLiee,s_CleTabLien,s_ChampLibLiee : String;

    VarIn : TStringList;                     // Récupére le paramètrage au lancement du menu. (cf TOF DPGRPTRAVAIL_TOF)

    TV_Width,                                // Largeur du composant TreeView à l'initialisation.
    MaxId,MaxPrgBar,ColumnToSort : integer;  // MaxId: Recupere la valeur max du n° Id table GRPDONNEES.

    Tb_Uti,               // Gestion des données de la table UTILISAT.
    Tb_GrpTrav,           // Gestion de la vue Hierarchique
    Tb_GrpDos,            // Gestion des affectations XXX (FILTREDONNEES)
    Tb_GrpUsr,            // Gestion des affectations UTILISATEURS
    Tb_Erase,             // Gestion des suppressions des affectations.
    Tb_InsOrUpd,          // Gestion des Transactions Insert ou Update.
    Tb_MultiSelect,
    Tb_ParamD : TOB;      // Gestion des données de la table paramdonnees.

    b_flg_Valid,b_DragnDrop : Boolean;// Confirmation d'enregistrement des modifs avant fermeture.

    // Tableau d'indexation de l'arbre...
    ArrTree : TTreeItemCodes ;
    // Tableau d'indexation de recherche d'un élément de la liste.
    ArrSearch : TSearchData;
    // Tableau d'élément de l'arborescence pour la suppression multiple.
    ArrNode,ArrNodeMulti : TArrayOfTreeNode;//array of TTreeNode;

    // Variable de Tri
    OrderSort : TOrdreTris;

    procedure DoDragDrop         (Sender, Source : TObject; X, Y : Integer);
    procedure DoDragOver         (Sender, Source : TObject; X, Y : Integer; State: TDragState; var Accept: Boolean);
    procedure DoEndDrag          (Sender, Target : TObject; X, Y : Integer);
    procedure DoStartDrag        (Sender : TObject; var DragObject: TDragObject);
    procedure DoMouseDown        (Sender : TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure DoMouseMove        (Sender : TObject; Shift: TShiftState; X, Y: Integer);
    procedure DoChange           (Sender : TObject; Node: TTreeNode);
    procedure DoCustomDrawItem   (Sender : TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoKeyDown          (Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure DoExpanding        (Sender : TObject; Node: TTreeNode; var AllowExpansion: Boolean);
    procedure DoBeforeExpandNode (noeud : TTreeNode);

    procedure DoClickPop_Menu    (Sender : TObject);
    procedure DoClickPop_SSMenu  (Sender : TObject);
    procedure DoClickPop_Supp    (Sender : TObject);
    procedure DoClickRdGrp       (Sender : TObject);
    procedure DoClickPop_Rename  (Sender : TObject);
    procedure DoClickPop_profil  (Sender : TObject);
    procedure DoClickPop_Rech    (Sender : TObject);
    procedure DoClickDragnDrop   (Sender : TObject);
    procedure DoClickAllDragnDrop(Sender : TObject);
    procedure DoDblClick         (Sender : TObject);
    procedure DoEnter            (Sender : TObject);
    procedure DoExit             (Sender : TObject);
    procedure DoSearchData       (Sender : TObject);
    procedure DoClickBInsert     (Sender : TObject);
    procedure DoClickBvoir       (Sender : TObject);
    procedure DoActiveMulAff     (Sender : TObject);
    procedure DoAnnuleMulAff     (Sender : TObject);

    // gestion de la fenêtre (type bloc note) de recherche.
    procedure LanceFenetreRecherche(ItemState : integer);
    procedure DoOnClick          (Sender : TObject);
    procedure DoOnCustomDrawItem (Sender: TCustomListView;Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure DoOnColumnClick    (Sender: TObject; Column: TListColumn);
    procedure DoOnCompare        (Sender: TObject; Item1, Item2: TListItem; Data: Integer; var Compare: Integer);

    procedure DoBeforeDragnDrop;

    procedure AjoutEnfant        (noeud : TTreeNode;bProfil : Boolean = TRUE);
    procedure MiseAJour_TOB      (noeud : TTreeNode;DragnDrop : Boolean = FALSE );
    procedure AutoriseAffectation(noeud : TTreeNode);
    procedure InitForm           (Fiche : TForm);

    procedure LectureArbre              (bLoading : Boolean = FALSE );
    procedure ReorganisationGrpDonnees  (bLoading : Boolean = FALSE );
    procedure RechercheNoeud            (Sender: TObject;St_donnee : string);
    procedure SuppAffectationDos        (Noeud : TTreeNode);
    procedure SuppAffectationUsr        (Noeud : TTreeNode);
    procedure ChargementTOB;
    procedure ChargementArbre;
    procedure GestionDesDroits;

    procedure RechercheNoeudEnfants (NoeudParent,Noeudenfant : TTreeNode);

    function RetourneCode       (St_chaine : string) : string;
    function RetourneProfil     (St_chaine : string) : string;
    function RetourneLibelle    (St_chaine : string ; bProfil : Boolean = FALSE) : string;
    function ExisteAffectation  (St_Code : string ; St_Donnee : string = '') : Boolean;

    procedure AddToArrTree   (Cod : String ; TreeItm : TTreeNode) ;
    function  GetItmFromCode (Cod : String ) : TTreeNode ;

    procedure AddToArrSearch   (ItemList : String ; IdColor : integer);
    function  GetIndexFromItem (ItemList : String) : Integer ;

    procedure AddToArrNode    (var ArrayNode: TArrayOfTreeNode;  Noeud : TTreeNode);
    procedure DoDeleteArrNode (Sender: TObject);
  end;

  procedure AffecteDosGrpConf (sNoDossier: string; bDelete : Boolean = FALSE);

Implementation
uses Grids,TntGrids,{$IFDEF DP}AnnOutils,{$ENDIF}
PROFIL_GRPTRAVAIL_TOF //mcd 16/07/07 sinon, n'existe pas dans le projet, alors que c'est obligatoire
,ed_tools,TntWideStrings,AGLinit;
////////////////////////////////////////////////////////////////////////////////
{procedure TOF_GRPTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GRPTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GRPTRAVAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GRPTRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_GRPTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;}
////////////////////////////////////////////////////////////////////////////////
{procedure TOF_GRPTRAVAIL.DoSelectAll(Sender: TObject);
var
  f_form : TForm;
begin
  //Non activé pour le moment, contrainte sur des gros volumes. KPMG table dossier
  //lors des tests 93 000 enreg.
  f_form := obj_Pnlst.InsideForm;

  while TFMul(f_form).BNext.Enabled do
    TFMul(f_form).BNextClick(nil);

  TFMul(f_form).bSelectAllClick(nil);
end;}
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.OnArgument (S : String ) ;
var
  s_param : string;
  f_form : TForm;
  IdImage : integer;
  arg : string ;
begin
  Inherited ;
  // Initialisation
  B_flg_Valid := TRUE;
  MaxId := 0;
  MaxPrgBar :=0;
  SetLength(ArrTree,0);
  VarIn := TStringList.Create;

  //Affichage de la barre de progression
  InitMoveProgressForm(nil,ecran.Caption,'Initialisation en cours, veuillez patienter...',2,FALSE,TRUE);
  TextProgressForm('Chargement des données');

  while s <> '' do
  begin
   {0 - FTD_CODE de la table FILTREDONNEES
    1 - Nature fiche multicritère
    2 - Nom fiche multicritère
    3 - Mode d'affectation Utilisateur Seulement
    4 - Intitulé du libelle pour option (par défaut "texte libre")
        sera alimenté dans la procédure ChargementTOB,
        par le libelle du nom de table DT_LIBELLE.DETABLES}
    s_param := READTOKENPipe(s,'|');
    VarIn.Add(s_param);
  end;

  //Chargements des TOBs.
  ChargementTOB;

  //init de la valeur Maxi pour la barre de progression
  if VarIn[3] <> 'X' then // Mode dossier
   MaxPrgBar := round(Tb_GrpDos.SommeSimple(S_tabprfx+'_MARK',[S_tabprfx+'_MARK'],['X'],TRUE))
  else                    // Mode Utilisateur seulement
   MaxPrgBar := Tb_GrpUsr.Detail.Count;

  // Panel d'affichage des fiches multicritère.
  obj_Pnlst := THPanel(GetControl('PNL_LST'));
  if (obj_Pnlst <> nil) and (VarIn[3] <> 'X') then  //mcd 22/05/07 ajout test UserOnly, dans ce cas, pas besoin de fiche
  begin
   {$IFDEF FILTREDONNEES}
    arg := 'FILTREDONNEES='+VarIn[0] ;//LM20071008
   {$ELSE}
    arg:='' ;
   {$ENDIF}
    AGLLanceFicheInside(VarIn[1], VarIn[2], '', '', arg, obj_Pnlst);//LM20071008
    f_form := obj_Pnlst.InsideForm;
    // Ctrl si existence de la fiche
    if f_form.Name = VarIn[2] then
      InitForm(f_form)
    else
      PGIINFO('Fiche "'+VarIn[2]+'" inexistante',ecran.Caption);
  end;

  Obj_pnldnd := THPanel(GetControl('PNL_DND'));
  Obj_redim  := TSplitter(GetControl('SPLIT'));
  if Obj_redim <> nil then
  begin
    Obj_redim.AutoSnap := FALSE;
    Obj_redim.MinSize := Obj_redim.Left;
  end;

  // POPUP MENU.
  if TMenuItem (GetControl ('MENU')) <> nil then
    TMenuItem (GetControl ('MENU')).OnClick       := DoClickPop_Menu;

  if TMenuItem (GetControl ('SOUS_MENU')) <> nil then
    TMenuItem (GetControl ('SOUS_MENU')).OnClick  := DoClickPop_SSMenu;

  if TMenuItem (GetControl ('SUP_OPT1')) <> nil then
    TMenuItem (GetControl ('SUP_OPT1')).OnClick   := DoClickPop_Supp;

  if TMenuItem (GetControl ('SUP_UNIQUE')) <> nil then
    TMenuItem (GetControl ('SUP_UNIQUE')).OnClick := DoClickPop_Supp;

  if TMenuItem (GetControl ('SUP_OPT2')) <> nil then
    TMenuItem (GetControl ('SUP_OPT2')).OnClick   := DoDeleteArrNode;

  if TMenuItem (GetControl ('SUP_ALL')) <> nil then
    TMenuItem (GetControl ('SUP_ALL')).OnClick    := DoDeleteArrNode;

  if TMenuItem (GetControl ('RENOMMER')) <> nil then
    TMenuItem (GetControl ('RENOMMER')).OnClick   := DoClickPop_Rename;

  if TMenuItem (GetControl ('RECHERCHER')) <> nil then
    TMenuItem (GetControl ('RECHERCHER')).OnClick := DoClickPop_Rech;

  if TMenuItem (GetControl ('PROFIL')) <> nil then
    TMenuItem (GetControl ('PROFIL')).OnClick     := DoClickPop_profil;

  if TMenuItem (GetControl ('MUL_OPT1')) <> nil then
    TMenuItem (GetControl ('MUL_OPT1')).OnClick     := DoActiveMulAff;

  if TMenuItem (GetControl ('MUL_OPT2')) <> nil then
    TMenuItem (GetControl ('MUL_OPT2')).OnClick     := DoAnnuleMulAff;

  // TREEVIEW.
  Obj_TreeVw := TTreeView(GetControl('TREE_HIER')) ;
  if Obj_TreeVw <> nil then
  begin
    TV_Width                    := Obj_TreeVw.Width;
    Obj_TreeVw.OnDragDrop       := DoDragDrop;
    Obj_TreeVw.OnDragOver       := DoDragOver;
    Obj_TreeVw.OnEndDrag        := DoEndDrag;
    Obj_TreeVw.OnMouseDown      := DoMouseDown;
    Obj_TreeVw.OnMouseMove      := DoMouseMove;
    Obj_TreeVw.OnChange         := DoChange;
    Obj_TreeVw.OnCustomDrawItem := DoCustomDrawItem;
    Obj_TreeVw.OnEnter          := DoEnter;
    Obj_TreeVw.OnExit           := DoExit;
    Obj_TreeVw.OnExpanding      := DoExpanding;
    Obj_TreeVw.OnKeyDown        := DoKeyDown;
    Obj_TreeVw.OnStartDrag      := DoStartDrag;

    // Chargement des icones standard PGI pour la TREEVIEW.
    if V_PGI.GraphList = nil then
      ChargeImageList;
    Obj_TreeVw.Images  := V_PGI.GraphList;
  end;

  // Radio bouton.
  Obj_RBtn_Dos := TRadioButton(GetControl('RDBTN_DOS'));
  if Obj_RBtn_Dos <> nil then
  begin
    Obj_RBtn_Dos.OnClick    := DoClickRdGrp;
    Obj_RBtn_Dos.Caption    := VarIn[4];
    Obj_RBtn_Dos.Hint       := 'Vue Hiérarchique par '+Obj_RBtn_Dos.Caption+' (Double-clic pour modifier le libellé)';
    Obj_RBtn_Dos.ShowHint   := TRUE;
    Obj_RBtn_Dos.OnDblClick := DoDblClick;
  end;

  Obj_RBtn_Usr := TRadioButton(GetControl('RDBTN_USR'));
  if Obj_RBtn_Usr <> nil then
    Obj_RBtn_Usr.OnClick := DoClickRdGrp;

  Obj_RBtn_All := TRadioButton(GetControl('RDBTN_ALL'));
  if Obj_RBtn_All <> nil then
    Obj_RBtn_All.OnClick := DoClickRdGrp;

  Obj_RBtn_None := TRadioButton(GetControl('RDBTN_NONE'));
  if Obj_RBtn_None <> nil then
    Obj_RBtn_None.OnClick := DoClickRdGrp;

  // Zone de Texte pour renommer l'intitulé du radio bouton.
  Obj_TxtLib := THEdit(GetControl('EDT_TXTLIBRE'));
  if Obj_TxtLib <> nil then
    Obj_TxtLib.OnExit := DoExit;

  // Boutons TToolbarButton97
  Obj_BtIns  := TToolbarButton97(GetControl('BINSERT'));
  if Obj_BtIns <> nil then
  begin
    Obj_BtIns.OnClick := DoClickPop_Menu;
    IdImage := V_PGI.GraphList.Add(Obj_BtIns.Glyph,Obj_BtIns.Glyph);
    V_PGI.GraphList.GetBitmap(IdImage,TMenuItem (GetControl ('MENU')).Bitmap);
  end;

  Obj_BtNIns := TToolbarButton97(GetControl('BNINSERT'));
  if Obj_BtNIns <> nil then
  begin
    Obj_BtNIns.OnClick := DoClickPop_SSMenu;
    IdImage := V_PGI.GraphList.Add(Obj_BtNIns.Glyph,Obj_BtNIns.Glyph);
    V_PGI.GraphList.GetBitmap(IdImage,TMenuItem (GetControl ('SOUS_MENU')).Bitmap);
  end;

  Obj_BtDel  := TToolbarButton97(GetControl('BDELETE'));
  if Obj_BtDel <> nil then
  begin
    //Obj_BtDel.OnClick := DoClickPop_Supp;
    IdImage := V_PGI.GraphList.Add(Obj_BtDel.Glyph,Obj_BtDel.Glyph);
    V_PGI.GraphList.GetBitmap(IdImage,TMenuItem (GetControl ('SUPPRIMER')).Bitmap);

    V_PGI.GraphList.GetBitmap(Ico_Delete,TMenuItem (GetControl ('SUP_OPT2')).Bitmap);
    V_PGI.GraphList.GetBitmap(Ico_Delete,TMenuItem (GetControl ('SUP_ALL')).Bitmap);
  end;

  Obj_BtRen  := TToolbarButton97(GetControl('BRENAME'));
  if Obj_BtRen <> nil then
  begin
    Obj_BtRen.OnClick := DoClickPop_Rename;
    IdImage := V_PGI.GraphList.Add(Obj_BtRen.Glyph,Obj_BtRen.Glyph);
    V_PGI.GraphList.GetBitmap(IdImage,TMenuItem (GetControl ('RENOMMER')).Bitmap);
  end;

  Obj_BtSearch := TToolbarButton97(GetControl('BSEARCH'));
  if Obj_BtSearch <> nil then
  begin
    Obj_BtSearch.OnClick := DoClickPop_Rech;
    IdImage := V_PGI.GraphList.Add(Obj_BtSearch.Glyph,Obj_BtSearch.Glyph);
    V_PGI.GraphList.GetBitmap(IdImage,TMenuItem (GetControl ('RECHERCHER')).Bitmap);
  end;

  if TMenuItem (GetControl ('MULTIAFFEC')) <> nil then
    V_PGI.GraphList.GetBitmap(Ico_GrpMulti,TMenuItem (GetControl ('MULTIAFFEC')).Bitmap);

  Obj_BtDnD := TToolbarButton97(GetControl('BDRAGNDROP'));
  if Obj_BtDnD <> nil then
    Obj_BtDnD.OnClick := DoClickDragnDrop;

  Obj_BtAllDnD := TToolbarButton97(GetControl('BALLDRAGNDROP'));
  if Obj_BtAllDnD <> nil then
    Obj_BtAllDnD.OnClick := DoClickAllDragnDrop;

  Obj_BtSearch_aff := TToolbarButton97(GetControl('BSEARCHAFF'));
  if Obj_BtSearch_aff <> nil then
    Obj_BtSearch_aff.OnClick := DoSearchData;

  // Affichage de la vue hierarchique sur la TreeView.
  ReorganisationGrpDonnees(TRUE);

  MaxValueProgressForm(MaxPrgBar);
  TextProgressForm('Chargement de l''arborescence');
  LectureArbre(TRUE);

  // Fin d'affichage de la barre de progression.
  FiniMoveProgressForm;

  // Désactive la 1ere option en fonction du mode de lancement. Si UserOnly = TRUE.
  if VarIn[3] = 'X' then
  begin
    Obj_RBtn_Dos.Enabled  := FALSE;
    Obj_RBtn_All.Enabled  := FALSE;
    Obj_RBtn_None.Enabled := FALSE;
    Obj_RBtn_Usr.Checked  := TRUE;
  end
  else
    TMenuItem (GetControl ('PROFIL')).Visible := FALSE;

  // Gestion des boutons TToolbarButton97. si TV est vide.
  if Obj_TreeVw.Items.Count = 0 then
  begin
    Obj_BtNIns.Enabled   := FALSE;
    Obj_BtDel.Enabled    := FALSE;
    Obj_BtRen.Enabled    := FALSE;
    Obj_BtSearch.Enabled := FALSE;
  end;

  GestionDesDroits;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.OnClose ;
begin
  Inherited ;
  // Sauvegarde si modifications apportées
  if (not B_flg_Valid) then
  begin
    if PGIAsk('Voulez-vous sauvegarder les modifications ?',ecran.caption)=mrYes then
      OnUpdate;
  end;

  //Libération
  FreeAndNil(LstView);
  FreeAndNil(WinFindResult);

  Tb_GrpTrav.Free;
  Tb_ParamD.Free;
  Tb_GrpDos.Free;
  Tb_GrpUsr.Free;
  Tb_Erase.Free;
  Tb_Uti.Free;
  Tb_InsOrUpd.Free;
  Tb_MultiSelect.Free;
  
  VarIn.Free;

  SetLength(ArrTree,0);
  SetLength(ArrSearch,0);
  SetLength(ArrNode,0);
  SetLength(ArrNodeMulti,0);
  Obj_Pnlst.InsideForm := nil;

  Obj_Pnlst.Free;
  Obj_TreeVw.Free;
  Obj_BtIns.Free;
  Obj_BtNIns.Free;
  Obj_BtDel.Free;
  Obj_BtRen.Free;
  Obj_BtSearch.Free;
  Obj_redim.Free;
  Obj_RBtn_Dos.Free;
  Obj_RBtn_Usr.Free;
  Obj_RBtn_All.Free;
  Obj_RBtn_None.Free;
  Obj_BtDnD.Free;
  Obj_BtAllDnD.Free;
  Obj_BtSearch_aff.Free;
  Obj_pnldnd.Free;
  Obj_TxtLib.Free;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_Menu(sender: Tobject);
var
  Noeud  : TTreeNode;
  tb_tmp : TOB;
  B_Exist : Boolean;
  S_nodeTxt,S_code,S_libelle,s_commentaire : string;
begin
  B_Exist := TRUE;
  s_commentaire := '';

  while B_Exist do
  begin
    // Appel fiche CREAT_GRPTRAVAIL.
    S_nodeTxt := AGLLanceFiche('DP','CREAT_GRPTRAVAIL', '', '', 'MENU');

    // Valeur de retour >> code,Libelle et libelle[code]
    S_code := RetourneCode(S_nodeTxt);

    // Création d'un Groupe de travail
    if S_code <> '' then
    begin
      // Recherche de l'existance du groupe de travail
      tb_tmp := tb_grptrav.FindFirst(['GRP_CODE'],[S_code],FALSE);
      if tb_tmp <> nil then
        PGIINFO('La codification "'+S_code+'" est déjà existante',ecran.caption)
      else
      begin
        B_Exist := FALSE;
        B_flg_Valid := FALSE;
        // Creation du groupe de travail
        Noeud := Obj_TreeVw.Items.Add(nil,S_nodeTxt);
        Noeud.ImageIndex    := Ico_Grp;
        Noeud.SelectedIndex := Ico_Grp;
        Noeud.HasChildren   := TRUE;
        Noeud.Selected      := TRUE;

        if not Obj_BtNIns.Enabled then
          Obj_BtNIns.Enabled := TRUE;
          
        if not Obj_BtDel.Enabled then
          Obj_BtDel.Enabled  := TRUE;

        // Mise à jour de la treeview.
        S_libelle := RetourneLibelle(S_nodeTxt);
        if Tb_GrpTrav.Detail.Count > 0 then
          s_commentaire := Tb_GrpTrav.Detail[0].GetString('GRP_COMMENTAIRE');

        inc(MaxId);

        tb_tmp := tob.create('GRPDONNEES',tb_grptrav,-1);
        tb_tmp.PutValue('GRP_ID',MaxId);
        tb_tmp.PutValue('GRP_IDPERE',-1);
        tb_tmp.PutValue('GRP_CODE',S_code);
        tb_tmp.PutValue('GRP_LIBELLE',S_libelle);
        tb_tmp.PutValue('GRP_NOM',VarIn[0]);
        tb_tmp.PutValue('GRP_COMMENTAIRE',s_commentaire);
        // Mise à jour du tableau
        AddToArrTree(S_code,Noeud);
      end;
    end
    else
      B_Exist := FALSE; // Annulation de création
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_SSMenu(sender: Tobject);
var
  id_Item,Idpere : Integer;
  Noeud,NoeudParent,NoeudUsr,NoeudLect : TTreeNode;
  tb_tmp : TOB;
  B_Exist : Boolean;
  S_comment,sCode,
  S_nodeTxt,S_code,S_libelle,S_profil,s_user : string;
begin
  B_Exist := TRUE;

  while B_Exist do
  begin
    // Appel fiche CREAT_GRPTRAVAIL.
    S_nodeTxt:= AGLLanceFiche('DP','CREAT_GRPTRAVAIL', '', '', 'SOUS_MENU');

    // valeur de retour >> code, libelle , libelle[code]
    S_code := RetourneCode(S_nodeTxt);
    S_libelle := RetourneLibelle(S_nodeTxt);

    // création d'un sous-groupe de travail.
    if S_code <> '' then
    begin
      //recherche de l'existance d'un sous-groupe de travail
      tb_tmp := tb_grptrav.FindFirst(['GRP_CODE'],[S_code],FALSE);
      if tb_tmp <> nil then
        PGIINFO('La codification "'+S_code+'" est déjà existante',ecran.caption)
      else
      begin
        B_Exist := FALSE;
        B_flg_Valid := FALSE;
        //Création du sous-groupe de travail
        id_Item := Obj_TreeVw.Selected.AbsoluteIndex;
        NoeudParent := Obj_TreeVw.Items[id_Item];

        //Ajout du nouveau sous menu
        Noeud := Obj_TreeVw.Items.Addchild(NoeudParent,S_nodeTxt);
        //Paramètre groupe dossier
        Noeud.ImageIndex    := Ico_Grp;
        Noeud.SelectedIndex := Ico_Grp;
        Noeud.Selected      := TRUE;

        //Récupére les infos liées au noeud parent.
        S_nodeTxt := NoeudParent.Text;
        sCode := RetourneCode(S_nodeTxt);

        tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[sCode],FALSE);
        if tb_tmp <> nil then
        begin
          inc(MaxId);
          Idpere := tb_tmp.GetInteger('GRP_ID');
          S_comment := tb_tmp.GetString('GRP_COMMENTAIRE');

          //Ajout nouvelle enreg.
          tb_tmp := TOB.create('GRPDONNEES',Tb_GrpTrav,-1);
          tb_tmp.PutValue('GRP_ID',MaxId);
          tb_tmp.PutValue('GRP_IDPERE',Idpere);
          tb_tmp.PutValue('GRP_CODE',S_code);
          tb_tmp.PutValue('GRP_LIBELLE',S_libelle);
          tb_tmp.PutValue('GRP_COMMENTAIRE',S_comment);
          tb_tmp.PutValue('GRP_NOM',VarIn[0]);
          //
          {tb_tmp := TOB.create('GRPDONNEES',Tb_InsOrUpd,-1);
          tb_tmp.PutValue('GRP_ID',MaxId);
          tb_tmp.PutValue('GRP_IDPERE',Idpere);
          tb_tmp.PutValue('GRP_CODE',S_code);
          tb_tmp.PutValue('GRP_LIBELLE',S_libelle);
          tb_tmp.PutValue('GRP_COMMENTAIRE',S_comment);
          tb_tmp.PutValue('GRP_NOM',VarIn[0]);}

          //Mise à jour du tableau
          AddToArrTree(S_code,Noeud);

          ReorganisationGrpDonnees;
        end;

        //si création d'un sous-menu en mode UTILISATEUR, replication
        //des affectations liées au noeud parent dans le nouveau sous-menu
        if Obj_RBtn_Usr.Checked then
        begin
          Obj_TreeVw.Items.BeginUpdate;
          NoeudLect := NoeudParent.getFirstChild;
          while (NoeudLect <> NoeudParent.getNextSibling) do
          begin
            s_user   := '';
            s_profil := '';

            if (NoeudLect.StateIndex = 1) and (NoeudLect <> Noeud) then
            begin
              s_user   := RetourneCode(NoeudLect.Text);
              S_profil := RetourneProfil(NoeudLect.Text);

              NoeudUsr := Obj_TreeVw.Items.AddChild(Noeud,NoeudLect.Text);
              NoeudUsr.StateIndex    := 1;
              NoeudUsr.ImageIndex    := Ico_Usr;
              NoeudUsr.SelectedIndex := Ico_Usr;
              // Màj Tob
              tb_tmp := TOB.create('LIENDONNEES',Tb_GrpUsr,-1);
              Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
              tb_tmp.PutValue('LND_GRPID',MaxId);
              tb_tmp.PutValue('LND_PROFILID',S_profil);
              tb_tmp.PutValue('LND_USERID',s_user);
              tb_tmp.PutValue('LND_NOM',VarIn[0]);
              // Màj
              tb_tmp := TOB.create('LIENDONNEES',Tb_InsOrUpd,-1);
              Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
              tb_tmp.PutValue('LND_GRPID',MaxId);
              tb_tmp.PutValue('LND_PROFILID',S_profil);
              tb_tmp.PutValue('LND_USERID',s_user);
              tb_tmp.PutValue('LND_NOM',VarIn[0]);

              NoeudLect := NoeudLect.GetNext;
            end
            else
            begin
              // Reduit le noeud pour passer outre ses enfants et passer au noeud suivant
              if NoeudLect.Expanded then
                NoeudLect.Collapse(true);
              NoeudLect := NoeudLect.GetNextVisible;
            end;

            if NoeudLect = nil then
              Break;
          end;
          Obj_TreeVw.Items.EndUpdate;
        end;
      end;
    end
    else
      B_Exist := FALSE; //Annulation de création
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_Supp(sender: Tobject);
var
  Noeud : TTreeNode;
  Tb_tmp,Tb_del : TOB;
  S_question,s_code : String;
  b_Delete : boolean;
  // Suppression d'un noeud et de tous ses enfants.
begin
  b_Delete := TRUE;
  Noeud    := Obj_TreeVw.Selected;

  case Noeud.StateIndex of
    0: begin
         S_question := 'Êtes-vous certain de vouloir retirer l''affectation "'+
                       Noeud.Text+'" du groupe de travail '+
                       Noeud.Parent.Text+' ?';
       end;
    1: begin
         S_question := 'Êtes-vous certain de vouloir retirer l''affectation "'+
                       Noeud.Text+'" du groupe de travail '+
                       Noeud.Parent.Text+' ?'+#13#10+
                       'Attention ! Pour un Profil/Utilisateur identique, le retrait s''éffectuera également sur les groupes de travail descendants.';
       end;
  else
    S_question := 'Êtes-vous certain de vouloir retirer le groupe de travail "'+Noeud.Text+'" ?';
  end;

  if Noeud <> nil then
  begin
    if PGIAsk(S_question,Ecran.Caption)=mrYes then
    begin
      case Noeud.StateIndex of
        0 : SuppAffectationDos(Noeud);
        1 : SuppAffectationUsr(Noeud);
      else
        // suppression d'un groupe de travail
        s_code := RetourneCode(Noeud.Text);
        //vérification s'il reste des affectations en cours.
        if not ExisteAffectation(s_code) then
        begin
          Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
          if Tb_tmp <> nil then
          begin
            Tb_del := tob.create('GRPDONNEES',tb_erase,-1);
            Tb_del.PutValue('GRP_ID',Tb_tmp.GetInteger('GRP_ID'));
            Tb_del.PutValue('GRP_IDPERE',Tb_tmp.GetInteger('GRP_IDPERE'));
            Tb_del.PutValue('GRP_CODE',Tb_tmp.GetString('GRP_CODE'));
            Tb_del.PutValue('GRP_NOM',Tb_tmp.GetString('GRP_NOM'));
            Tb_del.PutValue('GRP_LIBELLE',Tb_tmp.GetString('GRP_LIBELLE'));
            Tb_tmp.free;
          end;
        end
        else
        begin
          PGIInfo('Le retrait du groupe de travail "'+s_code +'" est impossible car il existe des affectations en cours.',Ecran.Caption);
          b_Delete := FALSE;
        end;
      end; // fin du case

      // Suppression du noeud dans la vue.
      if (Noeud <> nil) and (b_Delete) then
        Noeud.Delete;

      if Obj_TreeVw.Items.Count = 0 then
      begin
        Obj_BtNIns.Enabled   := FALSE;
        Obj_BtDel.Enabled    := FALSE;
        Obj_BtRen.Enabled    := FALSE;
        Obj_BtSearch.Enabled := FALSE;
      end;
      B_flg_Valid := FALSE;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickRdGrp(sender: Tobject);
var
  f_form : TForm;
  CC : TComponent;
  arg : string ;
begin
  //réinitialisation du tableau de recherche d'élément.
  SetLength(ArrSearch,0);

  //fermeture de la fenetre de recherche si active
  if assigned(WinFindResult) then
  begin
    FreeAndNil(LstView);
    FreeAndNil(WinFindResult);
  end;

  // Sélection du mode de travail via le radio bouton.
  if obj_Pnlst.InsideForm <> nil then
    obj_Pnlst.InsideForm := nil;

  if not obj_Pnlst.Visible then
  begin
    Obj_TreeVw.Align := alLeft;
    Obj_TreeVw.Width := TV_Width;

    Obj_redim.Visible := TRUE;
    Obj_redim.Align   := alLeft;

    Obj_Pnldnd.Visible := TRUE;
    Obj_Pnldnd.Align   := alLeft;

    obj_Pnlst.Visible := TRUE;
    obj_Pnlst.Align   := alClient;
  end;

  if Obj_RBtn_Dos.Checked then
  begin
   {$IFDEF FILTREDONNEES}
    arg := 'FILTREDONNEES='+VarIn[0] ;//LM20071008
   {$ELSE}
    arg:='' ;
   {$ENDIF}
    AGLLanceFicheInside(VarIn[1], VarIn[2], '', '', arg, obj_Pnlst); //LM20071008
    // Désactive le menu profil
    TMenuItem (GetControl ('PROFIL')).Visible := FALSE;

    f_form := obj_Pnlst.InsideForm;
    if f_form.Name = VarIn[2] then
      InitForm(f_form)
    else
      PGIINFO('Fiche "'+VarIn[2]+'" inexistante',ecran.Caption);
  end
  else if Obj_RBtn_Usr.Checked then
  begin
    AGLLanceFicheInside(Nat_Utilisat, Cod_Utilisat, '', '', '',obj_Pnlst);
    TMenuItem (GetControl ('PROFIL')).Visible := TRUE;

    f_form := obj_Pnlst.InsideForm;
    InitForm(f_form);
    AvertirTable('TTUSERGROUPE');

    if (v_pgi.Superviseur) then
    begin
      TFMul(f_form).Binsert.OnClick := DoClickBInsert;

      CC := TFMul(f_form).FindComponent('BVOIR');
      if CC <> nil then
        TToolbarButton97(CC).OnClick := DoClickBvoir;
    end;
  end
  else
  begin
    obj_Pnlst.Visible  := FALSE;
    Obj_Pnldnd.Visible := FALSE;
    Obj_redim.Visible  := FALSE;
    Obj_TreeVw.Align   := alClient;

    TMenuItem (GetControl ('PROFIL')).Visible := FALSE;
  end;
  //  Sauvegarde et recharge la vue....
  GestionDesDroits;
  ChargementArbre;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.ChargementTOB;
var
  S_Sql : String;
  Q : TQuery;
  tb_tmp : TOB;
begin
  // TOB
  tb_tmp := tob.create('TbTemporary',nil,-1);
  S_Sql  := 'SELECT * FROM FILTREDONNEES WHERE FTD_CODE = "'+VarIn[0]+'"';
  tb_tmp.LoadDetailFromSQL(S_Sql);

  s_tablien    := tb_tmp.detail[0].GetValue('FTD_TABLELIEN');
  s_tabliee    := tb_tmp.detail[0].GetValue('FTD_TABLELIEE');
  s_CleTabLiee := tb_tmp.detail[0].GetValue('FTD_CLETABLELIEE');
  s_CleTabLien := tb_tmp.detail[0].GetValue('FTD_CLETABLELIEN');

  S_tabprfx      := TableToPrefixe(s_tablien);
  s_ChampLibLiee := TableToPrefixe(s_tabliee)+'_LIBELLE';
  tb_tmp.free;

  //TOB réelle
  Tb_GrpTrav := TOB.Create('GRPDONNEES',nil,-1);
  Tb_ParamD  := TOB.Create('PARAMDONNEES',nil,-1);
  Tb_GrpDos  := TOB.Create(S_tablien,nil,-1);
  Tb_GrpUsr  := TOB.Create('LIENDONNEES',nil,-1);
  Tb_Uti     := TOB.Create('UTILISAT',nil,-1);

  //TOB virtuelle
  Tb_Erase       := TOB.Create('TbDelete',nil,-1);
  Tb_InsOrUpd    := TOB.Create('TbInsOrUpd',nil,-1);
  Tb_MultiSelect := TOB.Create('TbMultiSelect',nil,-1);

  //Test l'existence du champ dans la table.
  if not ChampLogiqueExiste(s_tabliee,s_ChampLibLiee) then
    s_ChampLibLiee := '';

  // Données + libellés liées à la table UTILISAT.
  S_Sql := 'SELECT US_UTILISATEUR,US_LIBELLE FROM UTILISAT ORDER BY US_UTILISATEUR';
  Tb_uti.LoadDetailFromSQL(s_sql);

  // Gestion des groupes de travail. mise en place de la TreeView.
  S_Sql := 'SELECT GRP_ID,GRP_IDPERE,GRP_CODE,'+
                  'GRP_LIBELLE,GRP_COMMENTAIRE,GRP_NOM '+
             'FROM GRPDONNEES WHERE GRP_NOM = "'+VarIn[0]+'" '+
            'ORDER BY GRP_IDPERE,GRP_LIBELLE';
  Tb_GrpTrav.LoadDetailDBFromSQL('GRPDONNEES',S_Sql,false,true);

  // Gestion des paramètres données.
  S_Sql := 'SELECT * FROM PARAMDONNEES WHERE PMD_GRPNOM = "'+VarIn[0]+'" '+
           'AND PMD_NOMTABLE = "'+s_tablien+'"';
  Tb_ParamD.LoadDetailDBFromSQL('PARAMDONNEES',S_Sql);

  // Valeur Max de Grp_id
  S_Sql := 'SELECT MAX(GRP_ID) MAXID FROM GRPDONNEES WHERE GRP_NOM = "'+VarIn[0]+'"';
  Q := OpenSQL(S_Sql,FALSE);
  if not Q.EOF then
    MaxId := Q.FindField('MAXID').AsInteger;
  Ferme(Q);

  // affectation du libelle. sur radio bouton Obj_RBtn_Dos.
  if Tb_ParamD.Detail.Count <> 0 then
    VarIn.Add(Tb_ParamD.Detail[0].GetString('PMD_LIBELLE'))
  else
    VarIn.Add(TableToLibelle(s_tablien));

  MoveCurProgressForm('Chargement des données');
end;
////////////////////////////////////////////////////////////////////////////////
// Sauvegarde de la vue hierarchique en base.
procedure TOF_GRPTRAVAIL.OnUpdate ;
var
  idx : Integer;
  StrMajCol : string;
  BooChampExist : Boolean;
begin
  Inherited ;
  StrMajCol     := TableToPrefixe(s_tabliee)+'_GROUPECONF';
  BooChampExist := ChampLogiqueExiste(s_tabliee,StrMajCol);

  InitMoveProgressForm(nil,ecran.Caption,'Enregistrement des modifications en cours, veuillez patienter...',6,FALSE,TRUE);

  // Suppression.
  for idx := 0 to Tb_Erase.Detail.Count - 1 do
  begin
    if (Tb_Erase.TOBNameExist(S_tablien)) and (BooChampExist) then
      AffecteDosGrpConf(Tb_Erase.Detail[idx].GetString(S_TabPrfx+'_NODOSSIER'),TRUE);

    Tb_Erase.Detail[idx].DeleteDB(FALSE);
    Application.ProcessMessages;
  end;
  MoveCurProgressForm;

  //Maj de la base de données.
  Tb_ParamD.InsertOrUpdateDB;

  Tb_GrpTrav.InsertOrUpdateDB;
  MoveCurProgressForm;

  // Maj du champ DOS_GROUPECONF. (Table Dossier)
  //Test l'existence du champ dans la table.
  for idx := 0 to Tb_InsOrUpd.Detail.Count -1 do
  begin
    if (Tb_InsOrUpd.TOBNameExist(S_tablien)) and (BooChampExist) then
      AffecteDosGrpConf(Tb_InsOrUpd.Detail[idx].GetString(S_TabPrfx+'_NODOSSIER'));

    Tb_InsOrUpd.Detail[idx].InsertOrUpdateDB;
    Application.ProcessMessages;
  end;

  FiniMoveProgressForm;
  B_flg_Valid := TRUE;
  PGIINFO('Modifications enregistrées',ecran.caption);
end ;
////////////////////////////////////////////////////////////////////////////////
// Affichage de la vue hierarchique.
procedure TOF_GRPTRAVAIL.LectureArbre(bLoading : Boolean);
var
  S_Dossier,S_User,S_Mark,s_libprof,
  s_code,s_profil,s_libelle : String;
  Noeud,NoeudChild : TTreeNode;
  Tb_tmp : TOB;
                         ////////////////////////
  procedure ReadGrpDos;
  var
    idx : integer;
  begin
    Obj_TreeVw.Items.BeginUpdate;
    for idx:=0 to Tb_GrpDos.Detail.Count -1 do
    begin
      s_code := Tb_GrpDos.Detail[idx].GetString(S_tabprfx+'_CODE');
      Noeud  := GetItmFromCode(s_code);

      S_Dossier := Tb_GrpDos.Detail[idx].GetString(s_CleTabLien);
      s_libelle := Tb_GrpDos.Detail[idx].GetString(S_tabprfx+'_DOSSLIB');
      S_Mark    := Tb_GrpDos.Detail[idx].GetString(S_tabprfx+'_MARK');

      if S_Mark = 'X' then
      begin
        S_Dossier := s_libelle+'['+S_Dossier+']';

        // Incremente la barre de progression de 1.
        if bLoading then
          MoveCurProgressForm('Chargement de l''arborescence');

        NoeudChild := Obj_TreeVw.Items.Addchild(Noeud,S_Dossier);
        // Init. propriétés du noeud.
        NoeudChild.StateIndex := 0;
        NoeudChild.ImageIndex := Ico_Dos;
        NoeudChild.SelectedIndex := Ico_Dos;
      end;
      Application.ProcessMessages;
    end;
    Obj_TreeVw.Items.EndUpdate;
  end;
                         ////////////////////////
  procedure ReadGrpUsr;
  var
    idx : integer;
  begin
    // Gestion des Utilisateurs.
    Obj_TreeVw.Items.BeginUpdate;
    for idx:=0 to Tb_GrpUsr.Detail.Count -1 do
    begin
      s_code := Tb_GrpUsr.Detail[idx].GetString('LND_CODE');
      Noeud  := GetItmFromCode(s_code);

      S_User   := Tb_GrpUsr.Detail[idx].GetString('LND_USERID');
      s_profil := Tb_GrpUsr.Detail[idx].GetString('LND_PROFILID');

      //Rattachement du libelle.
      Tb_tmp := AglTobRechercheOptimise(tb_uti,'US_UTILISATEUR',S_User);
      if Tb_tmp <> nil then
        S_User := Tb_tmp.GetString('US_LIBELLE')+'['+S_User+']'
      else
        S_User := '['+S_User+']';

      //Rattachement du profil
      if s_profil <> '' then
      begin
        s_libprof := RetourneLibelle(s_profil,TRUE);
        S_User := S_User+' ~ '+s_libprof+'['+s_profil+']';
      end;

      // Incremente la barre de progression de 1.
      if bLoading then
        MoveCurProgressForm('Chargement de l''arborescence');

      NoeudChild := Obj_TreeVw.Items.Addchild(Noeud,S_User);
      // Init. propriétés du noeud.
      NoeudChild.StateIndex    := 1;
      NoeudChild.ImageIndex    := Ico_Usr;
      NoeudChild.SelectedIndex := Ico_Usr;
      Application.ProcessMessages;
    end;
    Obj_TreeVw.Items.EndUpdate;
  end;
                         ////////////////////////
begin
  //Init.
  s_Search  := '';

  //Recherche des dossiers ou utilisateur rattachés aux noeuds.
  if Tb_Grptrav.Detail.Count > 0 then
  begin
    if (Obj_RBtn_Dos.Checked) then
      ReadGrpDos
    else if (Obj_RBtn_Usr.Checked) then
      ReadGrpUsr
    else if (Obj_RBtn_All.Checked) then
    begin
      Obj_TreeVw.Items.BeginUpdate;
      ReadGrpDos;
      ReadGrpUsr;
      Obj_TreeVw.Items.EndUpdate;
    end;
    //se positionne sur le 1er élément de la vue.
    Obj_TreeVw.SetFocus;
    Obj_TreeVw.Items[0].Selected := TRUE;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  Noeud : TTreeNode;
  B_autorisation : Boolean;
  Tb_tmp : TOB;
  CaseDragnDrop : integer; // 0 > affectation liste vers l'arborescence
                           // 1 > Deplacement groupe de travail
begin
  // Se produit quand l'utilisateur fait glisser un objet au-dessus d'un contrôle.
  // Ctrl et donne l'autorisation du drag & drop.
  B_autorisation := FALSE;
  CaseDragnDrop  := -1;

  {$IFDEF EAGLCLIENT}
    if (Source is THGrid) and (Sender is TTreeView) then
      CaseDragnDrop := 0
  {$ELSE}
    if (Source is THDBGrid) and (Sender is TTreeView) then
      CaseDragnDrop := 0
  {$ENDIF}
    else if (Source is TTreeView) and (Sender is TTreeView) then
      CaseDragnDrop := 1;

  case CaseDragnDrop of
    0 : begin
          Noeud := Obj_TreeVw.GetNodeAt(X,Y);
          if Noeud <> nil then
            AutoriseAffectation(Noeud);

          //recherche si il y a au moins un élément à affecter à un grp de travail.
          Tb_tmp := Tb_MultiSelect.FindFirst(['MUL_AUTORISE'],['X'],FALSE);
          if tb_tmp <> nil then
            B_autorisation := TRUE;

          if (Obj_TreeVw.Items.Count > 0) and (B_autorisation) then
            Accept := TRUE
          else
            Accept := FALSE;
        end;
    1 : begin
          // Gestion du défilement auto de la treeview lors
          // d'un drag n drop
          if (y < 15) then { scroll up }
          begin
            SendMessage(Obj_TreeVw.Handle, WM_VSCROLL, SB_LINEUP, 0);
            Obj_TreeVw.Refresh;
          end
          else if (Obj_TreeVw.Height - y < 15) then { scroll down }
          begin
            SendMessage(Obj_TreeVw.Handle, WM_VSCROLL, SB_LINEDOWN, 0);
            Obj_TreeVw.Refresh;
          end;


          // autorise un déplacement uniquement sur une groupe de travail
          noeud := Obj_TreeVw.GetNodeAt(X,Y);
          if (noeud <> nil) and (noeud.StateIndex = -1) and (noeud <> (Source as TTreeView).Selected) then
          begin
            AutoriseAffectation(Noeud);
            Tb_tmp := Tb_MultiSelect.FindFirst(['MUL_AUTORISE'],['X'],FALSE);
            if tb_tmp <> nil then
              B_autorisation := TRUE;

            if B_autorisation then
              Accept := true
            else
              Accept := false;
          end
          else
            Accept := false;
        end;
  else Accept := FALSE;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoDragDrop(Sender, Source: TObject; X,Y: Integer);
var
  Noeud,NoeudSrc,NoeudEnf : TTreeNode;
  CaseDragnDrop : integer; // 0 > affectation de liste vers l'arborescence
                           // 1 > déplacement dans l'arborescence
begin
  // Se produit quand l'utilisateur dépose un objet qu'il faisait glisser.
  CaseDragnDrop := -1;
  b_DragnDrop   := FALSE;

  {$IFDEF EAGLCLIENT}
    if (Sender is TTreeView) and (Source is THGrid) then
      CaseDragnDrop := 0
  {$ELSE}
    if (Sender is TTreeView) and (Source is THDBGrid) then
      CaseDragnDrop := 0
  {$ENDIF}
    else if (Sender is TTreeView) and (Source is TTreeView) then
      CaseDragnDrop := 1;

  case CaseDragnDrop of
    0 : begin
          // Ajout un nouvel item dans la vue.
          if (Obj_TreeVw.Items.count <> 0) then
          begin
            Noeud := Obj_TreeVw.GetNodeAt(X,Y); //noeud parent
            if Noeud <> nil then
            begin
              DoBeforeExpandNode(Noeud);
              AjoutEnfant(Noeud);
            end;
          end;
        end;
    1 : begin
          // Noeud source (Départ)
          NoeudSrc := Obj_TreeVw.Selected;
          // Noeud d'attache (Arrivée)
          Noeud    := Obj_TreeVw.GetNodeAt(X,Y);

          // Message utilisateur pour confirmation
          if PGIAsk('Êtes-vous certain de vouloir déplacer "'+NoeudSrc.Text+'" vers "'+Noeud.Text+'" ?',Ecran.Caption)=mrYes then
          begin
            if (Noeud <> nil) and (NoeudSrc <> nil) then
            begin
              b_DragnDrop := TRUE;
              // Déplacement d'une affectation vers un autre groupe de travail
              if (noeud.StateIndex = -1) and (NoeudSrc.StateIndex <> -1) then
              begin
                case NoeudSrc.StateIndex of
                  0: begin
                       NoeudEnf := Obj_TreeVw.Items.AddChild(noeud,noeudsrc.Text);
                       NoeudEnf.StateIndex    := 0;
                       NoeudEnf.ImageIndex    := Ico_Dos;
                       NoeudEnf.SelectedIndex := Ico_Dos;
                       NoeudEnf.Selected      := TRUE;
                       // Suppression du noeud de départ dans la TOB Tb_GRPDOS
                       SuppAffectationDos(NoeudSrc);
                       // Mise à jour de la Tob TB_GRPDOS
                       MiseAJour_TOB(noeud);
                       // Suppression du noeud de départ
                       NoeudSrc.Delete;
                     end;
                  1: begin
                       // Suppression du noeud de départ dans la TOB Tb_GRPUSR
                       SuppAffectationUsr(NoeudSrc);
                       AjoutEnfant(Noeud,FALSE);
                     end;
                end;
              end
              // Déplacement d'un groupe de travail (+ ses enfants) vers un
              // autre groupe de travail
              else if (noeud.StateIndex = -1) and (noeudsrc.StateIndex = -1) then
              begin
                NoeudEnf := Obj_TreeVw.Items.AddChild(noeud,noeudsrc.Text);
                NoeudEnf.StateIndex    := NoeudSrc.StateIndex;
                NoeudEnf.ImageIndex    := NoeudSrc.ImageIndex;
                NoeudEnf.SelectedIndex := NoeudSrc.SelectedIndex;
                // Recherche noeud enfant.
                if NoeudSrc.Count > 0 then
                  RechercheNoeudEnfants(NoeudSrc,NoeudEnf);
                // Suppression du noeud de départ
                NoeudSrc.Delete;
              end;
            end;
          end;
        end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoEndDrag(Sender, Target: TObject; X, Y: Integer);
var
  Noeud : TTreenode;
  CaseDragnDrop : integer;
begin
  // Se produit lorsque l'utilisateur arrête le glissement d'un objet,
  // soit en déposant l'objet, soit en annulant le glissement.
  CaseDragnDrop := -1;

  {$IFDEF EAGLCLIENT}
    if (Sender is THGrid) then
      CaseDragnDrop := 0
  {$ELSE}
    if (Sender is THDBGrid) then
      CaseDragnDrop := 0
  {$ENDIF}
    else if (Sender is TTreeView) then
      CaseDragnDrop := 1;

  case CaseDragnDrop of
    0 : begin
          if Target <> nil then
          begin
            Noeud := Obj_TreeVw.Selected;
            if Noeud <> nil then
              MiseAJour_TOB(Noeud);
          end;
        end;
    1 : begin
         exit; //provisoire
        end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Noeud : TTreenode;
begin
  if (sender is TTreeView) then
  begin
    Noeud := Obj_TreeVw.GetNodeAt(X,Y);
    if (Button = mbRight) then
    begin
      if v_pgi.Superviseur then
      begin
        if (Obj_TreeVw.Items.Count = 0) or (Noeud = nil) then
        begin
          TMenuItem (GetControl ('SOUS_MENU')).Enabled  := FALSE;
          TMenuItem (GetControl ('SUPPRIMER')).Enabled  := FALSE;
          TMenuItem (GetControl ('RENOMMER')).Enabled   := FALSE;
          TMenuItem (GetControl ('RECHERCHER')).Enabled := FALSE;
          TMenuItem (GetControl ('PROFIL')).Enabled     := FALSE;
          TMenuItem (GetControl ('MULTIAFFEC')).Enabled := FALSE;
        end
        else
        begin
          if (Noeud <> nil) and (Noeud.StateIndex = -1) then
          begin
            TMenuItem (GetControl ('SOUS_MENU')).Enabled  := TRUE;
            TMenuItem (GetControl ('RENOMMER')).Enabled   := TRUE;
            TMenuItem (GetControl ('MULTIAFFEC')).Enabled := TRUE;
            TMenuItem (GetControl ('RECHERCHER')).Enabled := FALSE;
            TMenuItem (GetControl ('PROFIL')).Enabled     := FALSE;
            // n'autorise pas la suppression d'un noeud parent
            // s'il existe des enfants.
            if Noeud.Count > 0 then
              TMenuItem (GetControl ('SUPPRIMER')).Enabled := FALSE
            else
              TMenuItem (GetControl ('SUPPRIMER')).Enabled := TRUE;

            //désactive la suppression multiple.
            TMenuItem (GetControl ('SUP_OPT2')).Enabled := FALSE;
          end
          else
          begin
            // n'autorise pas la création d'un sous groupe de travail
            // pour un noeud Dossier ou Utilisateur.
            TMenuItem (GetControl ('SOUS_MENU')).Enabled  := FALSE;
            TMenuItem (GetControl ('RENOMMER')).Enabled   := FALSE;
            TMenuItem (GetControl ('MULTIAFFEC')).Enabled := FALSE;
            TMenuItem (GetControl ('SUPPRIMER')).Enabled  := TRUE;
            TMenuItem (GetControl ('RECHERCHER')).Enabled := TRUE;
            TMenuItem (GetControl ('SUP_OPT2')).Enabled := TRUE;
            if Obj_RBtn_Usr.Checked then
              TMenuItem (GetControl ('PROFIL')).Enabled := TRUE;
          end;
        end;
      end
      else // not v_pgi.superviseur
      begin
        if (noeud <> nil) then
          TMenuItem (GetControl ('PROFIL')).Enabled := {$IFDEF DP}(JaiLeDroitAdminDossier) and{$ENDIF} (Obj_RBtn_Usr.Checked) and (noeud.StateIndex = 1)
        else
          TMenuItem (GetControl ('PROFIL')).Enabled := FALSE;
      end;
    end
    else if (Button = mbleft) then
    begin
      if (Noeud <> nil) and (Noeud.StateIndex = -1) then
        DoBeforeExpandNode(Noeud);
      Obj_TreeVw.Refresh;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoChange(Sender: TObject; Node: TTreeNode);
var
  Noeud : TTreenode;
begin
  // Valable en mode admin
  if v_pgi.Superviseur then
  begin
    Noeud := Obj_TreeVw.Selected;
    if (Noeud <> nil) then
    begin
      if (Noeud.StateIndex = -1) then
      begin
        Obj_BtNIns.Enabled   := TRUE;
        Obj_BtRen.Enabled    := TRUE;
        Obj_BtSearch.Enabled := FALSE;
        // n'autorise pas la suppression d'un noeud parent
        // s'il existe des enfants.
        if Noeud.Count > 0 then
          Obj_BtDel.Enabled := FALSE
        else
          Obj_BtDel.Enabled := TRUE;

        TMenuItem (GetControl ('SUP_ALL')).Enabled := FALSE;
      end
      else
      begin
        // n'autorise pas la création d'un sous groupe de travail
        // sur un noeud enfant.
        Obj_BtNIns.Enabled := FALSE;
        Obj_BtRen.Enabled := FALSE;
        Obj_BtDel.Enabled := TRUE;
        Obj_BtSearch.Enabled := TRUE;
        TMenuItem (GetControl ('SUP_ALL')).Enabled := TRUE;
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_Rename(sender: Tobject);
var
  id_Item : integer;
  Noeud : TTreeNode;
  tb_tmp : TOB;
  S_nodeTxt,s_code,s_libelle : string;
begin
  // modification du libellé (noeud données)
  id_Item := Obj_TreeVw.Selected.AbsoluteIndex;
  Noeud   := Obj_TreeVw.Items[id_Item];
  if (Noeud <> nil) and (Noeud.StateIndex = -1) then
  begin
    S_nodeTxt := AGLLanceFiche('DP','CREAT_GRPTRAVAIL', '', '',Noeud.Text);
    if S_nodeTxt <> '' then
    begin
      Noeud.Text := S_nodeTxt;
      // MiseAJour_Arbre;
      B_flg_Valid := FALSE;
      // Mise à jour de la TOB Tb_GrpTrav
      s_code := RetourneCode(Noeud.Text);
      s_libelle := RetourneLibelle(Noeud.Text);
      tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
      if tb_tmp <> nil then
        tb_tmp.PutValue('GRP_LIBELLE',s_libelle);
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_Rech(Sender: TObject);
var
  Noeud : TTreeNode;
begin
  SetLength(ArrSearch,0);
  // Noeud de départ.
  Noeud := Obj_TreeVw.Selected;
  s_Search := RetourneCode(Noeud.Text);
  AddToArrSearch(s_Search,0);

  RechercheNoeud(Sender,s_Search);

  LanceFenetreRecherche(Noeud.StateIndex);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoCustomDrawItem(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
 // Modifie le style du noeud. Suite menu recherche des occurrences
  case Node.StateIndex of
    0,1: begin
           sender.Canvas.Font.Color := clBlue;
           if (Node.Selected) and (sender.Focused) then
             sender.Canvas.Font.Color := clHighlightText
           else if (Node.Selected) and (not sender.Focused) then
             Node.Parent.Selected := true;
         end;
  else
    sender.Canvas.Font.Style := [];
    if (Node.Selected) and (not sender.Focused) then
    begin
      sender.Canvas.Font.Color := clHighlightText;
      sender.Canvas.Brush.Color := clHighlight;
    end;

    if (Node.StateIndex = -1) and (not v_pgi.Superviseur) and (node.ImageIndex = Ico_GrDenied) then
      sender.Canvas.Font.Color := clGrayText;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickPop_profil(Sender: TObject);
var
  idx : integer;
  Noeud : TTreeNode;
  s_enfant,s_state,s_retfich,
  s_code,S_Profil : string;
  tb_tmp,tb_tmp2 : TOB;
begin
  idx := Obj_TreeVw.Selected.AbsoluteIndex;
  Noeud := Obj_TreeVw.Items[idx];
  // utilisateur
  s_code := RetourneCode(noeud.parent.Text);
  s_enfant := RetourneCode(noeud.Text);
  S_Profil := RetourneProfil(noeud.Text);

  Tb_tmp := Tb_Grpusr.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,s_enfant,s_profil],FALSE);
  if Tb_tmp <> nil then
  begin
    s_profil  := Tb_tmp.GetString('LND_PROFILID');
    s_retfich := AGLLanceFiche('DP','PROFIL_GRPTRAVAIL', '', '', 'MODIFICATION|'+s_enfant+'|'+s_profil);
    s_state   := READTOKENPipe(s_retfich,'|');
    s_profil  := s_retfich;
    if s_state = 'mb_ok' then
    begin
      b_flg_Valid := FALSE;
      //recherche si profil déjà existant
      tb_tmp2 := Tb_Grpusr.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,s_enfant,s_profil],FALSE);
      if tb_tmp2 <> nil then
        PGIInfo('Le profil "'+s_profil+'" est déjà affecté à l''utilisateur "'+s_enfant+'"')
      else
      begin
        Tb_tmp.PutValue('LND_PROFILID',s_profil);
        if s_profil <> '' then
        begin
          if pos('~',Noeud.Text) > 0 then
            Noeud.text := RetourneLibelle(Noeud.Text)+'['+RetourneCode(Noeud.Text)+']'+' ~ '+RetourneLibelle(s_profil,TRUE)+'['+s_profil+']'
          else
            Noeud.text := Noeud.text+' ~ '+RetourneLibelle(s_profil,TRUE)+'['+s_profil+']';
        end
        else
          Noeud.text := RetourneLibelle(Noeud.Text)+'['+RetourneCode(Noeud.Text)+']';
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickDragnDrop(Sender: TObject);
var
  i : integer;

  procedure ExecuteAffectation(Noeud : TTreeNode);
  begin
    if noeud <> nil then
    begin
      //Autorisation des affectations
      AutoriseAffectation(noeud);
      //Ajout des items enfant
      AjoutEnfant(noeud);
      //Mise des tobs
      MiseAJour_TOB(noeud);
    end;
  end;

begin
  //récupére le ou les élements de la liste.
  DoBeforeDragnDrop;

  //Avertissement si aucun élément sélectionné.
  if Tb_MultiSelect.Detail.Count = 0 then
  begin
    PGIInfo('Veuillez sélectionner un ou plusieurs éléments de la liste',ecran.Caption);
    exit;
  end;

  // ctrl si multi sélection.
  if not Assigned(ArrNodeMulti) then
    ExecuteAffectation(Obj_TreeVw.Selected)
  else
  begin
    for i:=0 to High(ArrNodeMulti) do
    begin
      ExecuteAffectation(ArrNodeMulti[i]);
      ArrNodeMulti[i].ImageIndex := Ico_Grp;
      ArrNodeMulti[i].SelectedIndex := Ico_Grp;
    end;
    SetLength(ArrNodeMulti,0);
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoDblClick(Sender: TObject);
begin
  if (sender is TRadioButton) and (v_pgi.Superviseur) then
  begin
    if (sender as TRadioButton).Name = Obj_RBtn_Dos.Name then
    begin
      Obj_TxtLib.Text := Obj_RBtn_Dos.Caption;
      Obj_TxtLib.Visible := TRUE;
      Obj_TxtLib.SetFocus;
    end;
  end
  else
    //Désactive l'action du double click sur la liste.
    exit;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.AjoutEnfant(noeud: TTreeNode;bProfil : Boolean);
var
  NoeudChild,NoeudUsr,NoeudNext : TTreeNode;
  S_enfant,S_autorise,s_usr,s_retfich,s_profil,s_code,s_libelle : String;
  tb_tmp : tob;
  idx,GrpId : integer;
  b_AllPrf,b_NoPrf : boolean;
begin
  B_flg_Valid := FALSE;
  b_AllPrf    := FALSE;
  b_NoPrf     := FALSE;
  GrpId       := 0;

  for idx := 0 to Tb_MultiSelect.Detail.Count -1 do
  begin
    S_enfant    := Tb_MultiSelect.Detail[idx].GetString('MUL_SELECT');
    S_autorise  := Tb_MultiSelect.Detail[idx].GetString('MUL_AUTORISE');
    s_libelle   := Tb_MultiSelect.Detail[idx].GetString('MUL_LIBELLE');
    if not bProfil then
      s_profil    := Tb_MultiSelect.Detail[idx].GetString('MUL_PROFIL');
    s_usr       := S_enfant;
    if S_autorise = 'X' then
    begin
      // Recherche du libelle pour affichage enfant[libelle].
      if Obj_RBtn_Dos.Checked then
      begin
        S_enfant := s_libelle+'['+S_enfant+']';
        // Ajout nouveau noeud dans la vue.
        NoeudChild  := Obj_TreeVw.Items.Addchild(Noeud,S_enfant);
        // Init propriétés.
        NoeudChild.StateIndex    := 0;
        NoeudChild.ImageIndex    := Ico_Dos;
        NoeudChild.SelectedIndex := Ico_Dos;
        // Position sur le nouveau noeud.
        NoeudChild.Selected := TRUE;
      end
      else if Obj_RBtn_Usr.Checked then // GESTION UTILISATEUR
      begin
        //Profil utilisateur.
        if bProfil then
        begin
          if (not b_AllPrf) and (not b_NoPrf) then
          begin
            // s_retfich = Profil|AllProfil|NoProfil
            s_retfich := AGLLanceFiche('DP','PROFIL_GRPTRAVAIL', '', '', S_enfant);
            s_profil := READTOKENPipe(s_retfich,'|');

            if READTOKENPipe(s_retfich,'|') = 'X' then
              b_AllPrf := TRUE;

            if READTOKENPipe(s_retfich,'|') = 'X' then
              b_NoPrf := TRUE;
          end;
        end;

        // Récup du libelle correspondant au code utilisateur.
        Tb_tmp := AglTobRechercheOptimise(Tb_uti,'US_UTILISATEUR',S_enfant);
        if Tb_tmp <> nil then
          S_enfant := Tb_tmp.GetString('US_LIBELLE')+'['+S_enfant+']'
        else
          S_enfant := '['+S_enfant+']';

        // Affection du profil
        if s_profil <> '' then
          S_enfant := S_enfant+' ~ '+RetourneLibelle(s_profil,TRUE)+'['+s_profil+']';

        // Recherche du n° id  pour associer au nouvel enfant
        s_code := RetourneCode(noeud.Text);
        Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
        if Tb_tmp <> nil then
          GrpId := Tb_tmp.GetInteger('GRP_ID');

        // Replication de l'utilisateur dans les noeuds enfants.
        if noeud.HasChildren then
        begin
          //creation du premier noeud enfant
          Tb_tmp := Tb_Grpusr.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,S_usr,s_profil],FALSE);
          if Tb_tmp = nil then
          begin
            NoeudUsr := Obj_TreeVw.Items.AddChild(noeud,S_enfant);
            //Paramètres Utilisateur
            NoeudUsr.StateIndex    := 1;
            NoeudUsr.ImageIndex    := Ico_Usr;
            NoeudUsr.SelectedIndex := Ico_Usr;
            NoeudUsr.Selected      := TRUE;
            //maj de la TOB tb_grpusr
            Tb_tmp := tob.Create('LIENDONNEES',Tb_GrpUsr, -1);
            Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
            Tb_tmp.PutValue('LND_NOM',VarIn[0]);
            Tb_tmp.PutValue('LND_GRPID',GrpId);
            Tb_tmp.PutValue('LND_PROFILID',s_profil);
            Tb_tmp.PutValue('LND_USERID',s_usr);
            //
            Tb_tmp := tob.Create('LIENDONNEES',Tb_InsOrUpd, -1);
            Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
            Tb_tmp.PutValue('LND_NOM',VarIn[0]);
            Tb_tmp.PutValue('LND_GRPID',GrpId);
            Tb_tmp.PutValue('LND_PROFILID',s_profil);
            Tb_tmp.PutValue('LND_USERID',s_usr);
          end;
          //Duplication du noeud sur les groupes de données enfant.
          Noeudchild := noeud.GetFirstChild;
          NoeudNext  := noeud.getNextSibling;
          while (noeudchild <> NoeudNext) and (noeudchild.Level > noeud.Level) do
          begin
            if noeudchild.StateIndex = -1 then
            begin
              //ctrl existence user/profil
              s_code := RetourneCode(noeudchild.Text);

              //recherche du n° id
              Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
              if Tb_tmp <> nil then
                GrpId := Tb_tmp.GetInteger('GRP_ID');

              Tb_tmp := Tb_Grpusr.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,S_usr,s_profil],FALSE);
              if Tb_tmp = nil then
              begin
                NoeudUsr := Obj_TreeVw.Items.AddChild(noeudchild,S_enfant);
                //Paramètres Utilisateur
                NoeudUsr.StateIndex    := 1;
                NoeudUsr.ImageIndex    := Ico_Usr;
                NoeudUsr.SelectedIndex := Ico_Usr;
                //maj de la TOB tb_grpusr
                Tb_tmp := tob.Create('LIENDONNEES',Tb_GrpUsr, -1);
                Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
                Tb_tmp.PutValue('LND_NOM',VarIn[0]);
                Tb_tmp.PutValue('LND_GRPID',GrpId);
                Tb_tmp.PutValue('LND_PROFILID',s_profil);
                Tb_tmp.PutValue('LND_USERID',s_usr);
                //
                Tb_tmp := tob.Create('LIENDONNEES',Tb_InsOrUpd, -1);
                Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
                Tb_tmp.PutValue('LND_NOM',VarIn[0]);
                Tb_tmp.PutValue('LND_GRPID',GrpId);
                Tb_tmp.PutValue('LND_PROFILID',s_profil);
                Tb_tmp.PutValue('LND_USERID',s_usr);
              end;
            end;
            noeudchild := noeudchild.GetNext;
            if noeudchild = nil then
              break;
          end;
        end
        else
        begin
          //premier noeud enfant.
          NoeudChild := Obj_TreeVw.Items.AddChild(noeud,S_enfant);
          //Paramètre utilisateur.
          NoeudChild.StateIndex    := 1;
          NoeudChild.ImageIndex    := Ico_Usr;
          NoeudChild.SelectedIndex := Ico_Usr;
          NoeudChild.Selected      := TRUE;
          //maj de la TOB tb_grpusr
          s_code := RetourneCode(noeud.Text);
          Tb_tmp := tob.Create('LIENDONNEES',Tb_GrpUsr, -1);
          Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
          Tb_tmp.PutValue('LND_NOM',VarIn[0]);
          Tb_tmp.PutValue('LND_GRPID',GrpId);
          Tb_tmp.PutValue('LND_PROFILID',s_profil);
          Tb_tmp.PutValue('LND_USERID',s_usr);
          //
          Tb_tmp := tob.Create('LIENDONNEES',Tb_InsOrUpd, -1);
          Tb_tmp.AddChampSupValeur('LND_CODE',s_code);
          Tb_tmp.PutValue('LND_NOM',VarIn[0]);
          Tb_tmp.PutValue('LND_GRPID',GrpId);
          Tb_tmp.PutValue('LND_PROFILID',s_profil);
          Tb_tmp.PutValue('LND_USERID',s_usr);
        end;
      end;
    end;
  end; // fin du for.
  Obj_TreeVw.Refresh;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.MiseAJour_TOB(noeud: TTreeNode;DragnDrop : Boolean);
var
  S_enfant,S_autorise,s_parent,s_code,s_cod,s_libelle : String;
  NoeudParent : TTreenode;
  Tb_tmp : TOB;
  idx,IdMaster,IdSlave : integer;
begin
  IdMaster := 0;
  IdSlave  := 0;

  if DragnDrop then
    s_parent  := Noeud.parent.Text
  else
    s_parent  := Noeud.Text;

  s_code := RetourneCode(s_parent);
  //recherche du n° ID associé au code noeud.
  Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
  if Tb_tmp <> nil then
    IdMaster := Tb_tmp.GetInteger('GRP_ID');

  for idx := 0 to Tb_MultiSelect.Detail.Count -1 do
  begin
    S_enfant   := Tb_MultiSelect.Detail[idx].GetString('MUL_SELECT');
    S_autorise := Tb_MultiSelect.Detail[idx].GetString('MUL_AUTORISE');
    s_libelle  := Tb_MultiSelect.Detail[idx].GetString('MUL_LIBELLE');

    if S_autorise = 'X' then
    begin
      if Obj_RBtn_Dos.Checked then
      begin
        Tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE',s_CleTabLien],[s_code,S_enfant],FALSE);
        if Tb_tmp = nil then
        begin
          Tb_tmp := tob.Create(S_tablien,Tb_GrpDos, -1);
          Tb_tmp.AddChampSupValeur(S_tabprfx+'_CODE',s_code);
          Tb_tmp.AddChampSupValeur(S_tabprfx+'_DOSSLIB',s_libelle);
          Tb_tmp.PutValue(S_tabprfx+'_NOM',VarIn[0]);
          Tb_tmp.PutValue(S_tabprfx+'_GRPID',IdMaster);
          Tb_tmp.PutValue(S_tabprfx+'_MARK','X');
          Tb_tmp.PutValue(s_CleTabLien,S_enfant);
          //
          Tb_tmp := tob.Create(S_tablien,Tb_InsOrUpd, -1);
          Tb_tmp.AddChampSupValeur(S_tabprfx+'_CODE',s_code);
          Tb_tmp.AddChampSupValeur(S_tabprfx+'_DOSSLIB',s_libelle);
          Tb_tmp.PutValue(S_tabprfx+'_NOM',VarIn[0]);
          Tb_tmp.PutValue(S_tabprfx+'_GRPID',IdMaster);
          Tb_tmp.PutValue(S_tabprfx+'_MARK','X');
          Tb_tmp.PutValue(s_CleTabLien,S_enfant);
        end;
        // Identifie les noeuds parents pour remonter le dossier.
        if DragnDrop then
          NoeudParent := (Noeud.Parent).Parent
        else
          NoeudParent := Noeud.Parent;

        while NoeudParent<>nil do
        begin
          s_cod := RetourneCode(NoeudParent.Text);
          Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_cod],FALSE);
          if Tb_tmp <> nil then
            IdSlave := Tb_tmp.GetInteger('GRP_ID');

          Tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE',s_CleTabLien],[s_cod,S_enfant],FALSE);
          if Tb_tmp = nil then
          begin
            Tb_tmp := tob.Create(S_tablien,Tb_GrpDos, -1);
            Tb_tmp.AddChampSupValeur(S_tabprfx+'_CODE',s_cod);
            Tb_tmp.AddChampSupValeur(S_tabprfx+'_DOSSLIB',s_libelle);
            Tb_tmp.PutValue(S_tabprfx+'_NOM',VarIn[0]);
            Tb_tmp.PutValue(S_tabprfx+'_GRPID',IdSlave);
            Tb_tmp.PutValue(S_tabprfx+'_MARK','-');
            Tb_tmp.PutValue(s_CleTabLien,S_enfant);
            //
            Tb_tmp := tob.Create(S_tablien,Tb_InsOrUpd, -1);
            Tb_tmp.AddChampSupValeur(S_tabprfx+'_CODE',s_cod);
            Tb_tmp.AddChampSupValeur(S_tabprfx+'_DOSSLIB',s_libelle);
            Tb_tmp.PutValue(S_tabprfx+'_NOM',VarIn[0]);
            Tb_tmp.PutValue(S_tabprfx+'_GRPID',IdSlave);
            Tb_tmp.PutValue(S_tabprfx+'_MARK','-');
            Tb_tmp.PutValue(s_CleTabLien,S_enfant);
          end;
          NoeudParent := NoeudParent.parent;
        end;
      end;
    end;
  end; //fin du for
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.AutoriseAffectation(noeud: TTreeNode);
var
//  NoeudParent : TTreeNode;
  S_enfant,s_code : String;
  idx : integer;
  Tb_tmp : TOB;
begin
  // Ctrl et donne l'autorisation du drag & drop.
  for idx := 0 to Tb_MultiSelect.Detail.Count -1 do
  begin
    S_enfant := Tb_MultiSelect.Detail[idx].GetString('MUL_SELECT');

    if Noeud <> nil then
    begin
      S_code := RetourneCode(Noeud.Text);
      if Obj_RBtn_Dos.Checked then
      begin
        // Vérification de l'élément si déjà présent dans le groupe de travail.
        Tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE',s_CleTabLien],[S_code, S_enfant],FALSE);
        if ((Tb_tmp = nil) and (not ExisteAffectation(S_code,S_enfant))) and (Noeud.StateIndex = -1) then
          Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','X')
        else
          Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','-');
      end
      // GESTION UTILISATEUR
      else if Obj_RBtn_Usr.Checked then
      begin
        if v_pgi.NumVersionSoc < 844 then
        begin
          Tb_tmp := Tb_Grpusr.FindFirst(['LND_CODE','LND_USERID'],[S_code, S_enfant],FALSE);
          if ((Tb_tmp = nil) and (not ExisteAffectation(S_code,S_enfant))) and (Noeud.StateIndex = -1) then
            Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','X')
          else
            Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','-');
        end
        else
        begin
          // A partir de la version 845 possibilité d'affecter n fois le même
          // utilisateur à un groupe de travail pour un profil différent.
          // Ce ctrl s'effectue dans la procedure MiseAJour_TOB.
          if v_pgi.Superviseur then
          begin
            if (Noeud.StateIndex = -1) then
              Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','X')
            else
              Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','-');
          end
          else
          begin
            if (Noeud.StateIndex = -1) and (Noeud.ImageIndex = Ico_Grp) and (s_enfant <> v_pgi.User) then
              Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','X')
            else
              Tb_MultiSelect.Detail[idx].PutValue('MUL_AUTORISE','-');
          end;
        end;
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoEnter(Sender: TObject);
begin
  // Rafraichit l'affichage du composant TreeView lorsqu'il reçoit son focus.
  Obj_TreeVw.Refresh;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.InitForm(Fiche: TForm);
begin
  // Initialisation des différentes actions après ouverture de la fiche multi-critère.
  TFMul(Fiche).BAnnuler.Visible   := FALSE;
  TFMul(Fiche).BOuvrir.Visible    := FALSE;
  TFMul(Fiche).bSelectAll.Visible := TRUE;

  Ecran.OnKeyDown                 := DoKeyDown;

  // Action de la liste
  TFMul(Fiche).FListe.OnMouseDown := DoMouseDown;
  TFMul(Fiche).FListe.OnMouseMove := DoMouseMove;
  TFMul(Fiche).FListe.OnEndDrag   := DoEndDrag;
  TFMul(Fiche).FListe.OnDblClick  := DoDblClick;
  TFMul(Fiche).FListe.OnStartDrag := DoStartDrag;

  {$IFDEF EAGLCLIENT}
    TFMul(Fiche).FListe.MultiSelect := TRUE;
  {$ELSE}
    TFMul(Fiche).FListe.MultiSelection := TRUE;
  {$ENDIF}
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoStartDrag(Sender: TObject;
  var DragObject: TDragObject);
var
  noeud : TTreeNode;
  s_nodossier,s_libdossier,s_profil : string;
  Tb_tmp : TOB;
begin
  if (Sender is TTreeView) then
  begin
    Tb_MultiSelect.ClearDetail;
    noeud := (Sender as TTreeView).Selected;
    if (noeud <> nil) and (Noeud.StateIndex <> -1) then
    begin
      s_nodossier := RetourneCode(noeud.Text);
      s_libdossier:= RetourneLibelle(noeud.Text);
      s_profil    := RetourneProfil(noeud.Text);

      Tb_tmp := tob.create('',Tb_MultiSelect,-1);
      Tb_tmp.AddChampSupValeur('MUL_SELECT',s_nodossier);
      Tb_tmp.AddChampSupValeur('MUL_LIBELLE',s_libdossier);
      Tb_tmp.AddChampSupValeur('MUL_PROFIL',s_profil);
      Tb_tmp.AddChampSupValeur('MUL_AUTORISE','X');
    end;
  end
  else
    DoBeforeDragnDrop;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Fecran : Tform;
  Noeud : TTreeNode;
  i : integer;

  Procedure SetImageNode(Node : TTreeNode; Ico : integer);
  begin
    if Node.StateIndex = -1 then
    begin
      Node.ImageIndex    := Ico;
      Node.SelectedIndex := Ico;
    end
    else
    begin
      if Node.ImageIndex = Ico then
      begin
        Node.ImageIndex    := Ico_Delete;
        Node.SelectedIndex := Ico_Delete;
        AddToArrNode(ArrNode,Node);
      end
      else
      // Retire la Demande de suppression
      begin
        Node.ImageIndex    := Ico;
        Node.SelectedIndex := Ico;
      end;
    end;
  end;

begin
  if (sender is TTreeView) then
  begin
    case key of
      // Barre d'espace sur le groupe de travail. sélection ou déselection des
      //affectations pour retrait.
      VK_SPACE : begin
                   Noeud := Obj_TreeVw.Selected;
                   if Noeud <> nil then
                   begin
                     case Noeud.StateIndex of
                     -1 : begin
                            for i:=0 to noeud.Count-1 do
                            begin
                              case noeud.Item[i].StateIndex of
                                0 : SetImageNode(noeud.Item[i],Ico_Dos);
                                1 : SetImageNode(noeud.Item[i],Ico_Usr);
                              end;
                            end;
                          end;
                      0 : SetImageNode(Noeud,Ico_Dos);
                      1 : SetImageNode(Noeud,Ico_Usr);
                     end;
                   end;
                 end;
      // Ctrl + Alt + X = Sélection du groupe de travail pour une multi affectation
      Ord('X') : begin
                   if (Shift = [ssCtrl,ssAlt]) then
                   begin
                     Noeud := Obj_TreeVw.Selected;
                     if (Noeud <> nil) and (Noeud.StateIndex = -1) then
                     begin
                       SetImageNode(noeud,Ico_GrpMulti);
                       AddToArrNode(ArrNodeMulti,noeud);
                     end;
                   end;
                 end;
      // Ctrl + Alt + A = Désélection du groupe de travail pour une multi affectation
      Ord('A') : begin
                   if (Shift = [ssCtrl,ssAlt]) then
                   begin
                     Noeud := Obj_TreeVw.Selected;
                     if (Noeud <> nil) and (Noeud.StateIndex = -1) then
                       SetImageNode(noeud,Ico_Grp);
                   end;
                 end;
    end;
  end
  else
  begin
    Fecran := obj_Pnlst.InsideForm;
    case key of
      // Critéres de recherche
      VK_F9 : TFMul(Fecran).BCherche.Click;
      {$IFDEF EAGLCLIENT}
        // Exclusivement mode Cwas : Fonction PREMIER enreg d'une liste
        VK_PRIOR : TFMul(Fecran).BPrev.Click;
        // Exclusivement mode Cwas : Fonction DERNIER enreg d'une liste
        VK_NEXT : TFMul(Fecran).BNext.Click;
      {$ENDIF}
      // sélection ou déselection de tous les enreg d'une liste
      Ord('A') : if (Shift=[ssCtrl]) and (Obj_RBtn_Dos.Checked or Obj_RBtn_Usr.Checked)
                   then TFMul(Fecran).bSelectAll.Click; //Obj_BtDesel.Click;
      // Création nouveau Menu
      Ord('N') : if (Shift=[ssCtrl]) and (v_pgi.Superviseur) then
                   Obj_BtIns.Click;
      // Recherche des occurences d'un élément dans la vue de l'arborescence
      Ord('F') : begin
                   if (Shift=[ssCtrl]) and (v_pgi.Superviseur) then
                   begin
                     if Obj_BtSearch.Enabled then
                       Obj_BtSearch.Click;
                   end;
                 end;
      // Suppression dans la vue de l'arborescence
      VK_DELETE : begin
                    if (Shift=[ssCtrl]) and (v_pgi.Superviseur) then
                    begin
                      if Obj_BtDel.Enabled then
                        DoClickPop_Supp(nil);
                    end;
                  end;
      // Validation des modifications
      VK_F10 : OnUpdate;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoSearchData(Sender: TObject);
var
  idx,State : integer;
  Fecran : TForm;
  s_ColValue,s_Value : string;

  {$IFDEF EAGLCLIENT}
    HGrid : THGrid;
  {$ELSE}
    HGrid : THDBGrid;
  {$ENDIF}
begin
  // Recherche si un élement de la liste est affecté dans un groupe de travail
  Fecran := obj_Pnlst.InsideForm;
  HGrid  := TFMul(Fecran).FListe;
  SetLength(ArrSearch,0);
  S_Search := '';
  State    := 0;

  // Initialisation des noms de colonnes
  if Obj_RBtn_Dos.Checked then
  begin
    s_ColValue    := s_CleTabLiee;
    State := 0;
  end
  else if Obj_RBtn_Usr.Checked then
  begin
    s_ColValue    := 'US_UTILISATEUR';
    State := 1;
  end;

  if HGrid.nbSelected > 1 then
  begin
    // Récupére la sélection de la liste
    for idx :=0 to HGrid.nbSelected -1 do
    begin
      HGrid.GotoLeBookmark(idx);

      {$IFDEF EAGLCLIENT}
        TFMul(Fecran).Q.TQ.Seek(HGrid.Row-1);
      {$ENDIF}

      s_Value := TFMul(Fecran).Q.FindField(s_ColValue).AsString ;
      AddToArrSearch(s_Value,0);
      RechercheNoeud(sender,s_Value);
    end;
  end
  else
  begin
    {$IFDEF EAGLCLIENT}
      TFMul(Fecran).Q.TQ.Seek(HGrid.Row-1);
    {$ENDIF}

    s_Value := TFMul(Fecran).Q.FindField(s_ColValue).AsString ;
    AddToArrSearch(s_Value,0);
    RechercheNoeud(sender,s_Value);
  end;

  LanceFenetreRecherche(State);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.ReorganisationGrpDonnees(bLoading : Boolean);
var
  TobId,idx,idxpere,i,PosId : Integer;
  Tb_tmp,TobUsr : TOB;
  Noeud,NoeudChild,NoeudParent : TTreeNode;
  S_code,S_libelle,S_nodeTxt,S_Sql : string;

  function DroitAcces(grpid : integer; code, utilisat : string) : boolean;
  var
    TobPtr : TOB;
  begin
    Result := FALSE;
    TobPtr := TobUsr.FindFirst(['LND_GRPID','LND_USERID','LND_CODE'],[grpid,utilisat,code],FALSE);
    if TobPtr <> nil then
      Result := TRUE;
  end;

begin
  Noeud := nil;
  if not v_pgi.Superviseur then
  begin
    TobUsr := Tob.Create('VirtualTob',nil,-1);
    S_Sql :=    'SELECT LND_GRPID,LND_USERID,GRP_CODE "LND_CODE" '+
                  'FROM LIENDONNEES '+
             'LEFT JOIN GRPDONNEES ON LIENDONNEES.LND_GRPID = GRPDONNEES.GRP_ID '+
                                 'AND LIENDONNEES.LND_NOM   = GRPDONNEES.GRP_NOM '+
                 'WHERE LND_USERID  = "'+v_pgi.User+'"';
    TobUsr.LoadDetailFromSQL(S_Sql);
  end;

  Obj_TreeVw.Items.BeginUpdate;
  // Mise en place de l'arborescence.
  for i:=0 to Tb_Grptrav.Detail.Count -1 do
  begin
    idx     := Tb_Grptrav.Detail[i].GetInteger('GRP_ID');
    idxpere := Tb_Grptrav.Detail[i].GetInteger('GRP_IDPERE');

    if bLoading then
    begin
      S_code    := Tb_Grptrav.Detail[i].GetString('GRP_CODE');
      S_libelle := Tb_Grptrav.Detail[i].GetString('GRP_LIBELLE');
      S_nodeTxt := S_libelle+'['+S_code+']';

      //recherche si relation coherente entre idx et idxpere.
      Tb_tmp := Tb_Grptrav.FindFirst(['GRP_ID'],[idxpere],FALSE);
      if Tb_tmp = nil then
      begin
        noeud := Obj_TreeVw.Items.Add(nil,S_nodeTxt);
        noeud.HasChildren := TRUE;
      end
      else
        noeud := Obj_TreeVw.Items[i];

      AddToArrTree(S_code,Noeud);
    end;

    TobId := i;

    Tb_tmp := Tb_Grptrav.FindFirst(['GRP_IDPERE'],[idx],FALSE);
    while (Tb_tmp <> nil) do
    begin
      if bLoading then
      begin
        S_code    := Tb_tmp.GetString('GRP_CODE');
        S_libelle := Tb_tmp.GetString('GRP_LIBELLE');
        S_nodeTxt := S_libelle+'['+S_code+']';
        NoeudChild := Obj_TreeVw.Items.AddChild(noeud,S_nodeTxt);
        NoeudChild.HasChildren := TRUE;

        AddToArrTree(S_code,NoeudChild);
        if v_pgi.Superviseur then
        begin
          NoeudChild.ImageIndex    := Ico_Grp;
          NoeudChild.SelectedIndex := Ico_Grp;
        end
        else
        begin
          if not DroitAcces(idx,S_code,v_pgi.User) then
          begin
            NoeudChild.ImageIndex    := Ico_GrDenied;
            NoeudChild.SelectedIndex := Ico_GrDenied;
          end
          else
          begin
            NoeudChild.ImageIndex    := Ico_Grp;
            NoeudChild.SelectedIndex := Ico_Grp;
            //recherche des parents
            NoeudParent := noeud.Parent;
            while Noeudparent <> nil do
            begin
              Noeudparent.ImageIndex    := Ico_Grp;
              Noeudparent.SelectedIndex := Ico_Grp;
              NoeudParent := noeudparent.Parent;
            end;
          end;
        end;
      end;

      PosId := Tb_Grptrav.Detail.IndexOf(Tb_tmp);
      inc(TobId);
      Tb_Grptrav.Detail.Insert(TobId,Tb_tmp);
      Tb_Grptrav.Detail.Delete(PosId+1);
      Application.ProcessMessages;
      Tb_tmp := Tb_Grptrav.FindNext(['GRP_IDPERE'],[idx],FALSE);
    end;
    Application.ProcessMessages;

    if bLoading then
    begin
      if v_pgi.Superviseur then
      begin
        Noeud.ImageIndex    := Ico_Grp;
        Noeud.SelectedIndex := Ico_Grp;
      end
      else
      begin
        if not DroitAcces(idx,S_code,v_pgi.User) then
        begin
          Noeud.ImageIndex    := Ico_GrDenied;
          Noeud.SelectedIndex := Ico_GrDenied;
        end
        else
        begin
          Noeud.ImageIndex    := Ico_Grp;
          Noeud.SelectedIndex := Ico_Grp;
          NoeudParent := noeud.Parent;
          while Noeudparent <> nil do
          begin
            Noeudparent.ImageIndex    := Ico_Grp;
            Noeudparent.SelectedIndex := Ico_Grp;
            NoeudParent := noeudparent.Parent;
          end;
        end;
      end;
    end;
  end;
  Obj_TreeVw.Items.EndUpdate;
  if not v_pgi.Superviseur then
    FreeAndNil(TobUsr);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoExit(Sender: TObject);
var
  tb_tmp : TOB;
begin
  if (sender is TTreeView) then
    Obj_TreeVw.Refresh
  else if (sender is THEdit) then
  begin
    // Modif du libelle sur le radio bouton Obj_RBtn_Dos
    if Obj_TxtLib.Text <> '' then
    begin
      Obj_RBtn_Dos.Caption := Obj_TxtLib.Text;
      // On cache la zone de saisie.
      Obj_TxtLib.Visible := FALSE;
      // Maj de la Tob.
      Tb_ParamD.ClearDetail;
      tb_tmp := TOB.Create('PARAMDONNEES',Tb_ParamD,-1);
      tb_tmp.PutValue('PMD_GRPNOM',VarIn[0]);
      tb_tmp.PutValue('PMD_NOMTABLE',S_Tablien);
      tb_tmp.PutValue('PMD_SELECT','');
      tb_tmp.PutValue('PMD_ORDERBY','');
      tb_tmp.PutValue('PMD_LIBELLE',Obj_RBtn_Dos.Caption);
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.ChargementArbre;
begin
  Obj_TreeVw.Cursor := crHourGlass;

  InitMoveProgressForm(nil,ecran.Caption,'Chargement de l''arborescence, veuillez patienter...',2,FALSE,TRUE);

  SetLength(ArrTree,0);
  SetLength(ArrNode,0);
  SetLength(ArrNodeMulti,0);
  Obj_TreeVw.Items.Clear;

  ReorganisationGrpDonnees(TRUE);
  MoveCurProgressForm;

  LectureArbre;
  MoveCurProgressForm;

  Obj_TreeVw.Refresh;
  FiniMoveProgressForm;

  Obj_TreeVw.Cursor := crDefault;
end;
////////////////////////////////////////////////////////////////////////////////
function TOF_GRPTRAVAIL.RetourneCode(St_chaine: string): string;
begin
  //retourne le code entre []
  result := Copy(st_chaine,pos('[',st_chaine)+1,pos(']',st_chaine)-pos('[',st_chaine)-1);
end;
////////////////////////////////////////////////////////////////////////////////
function TOF_GRPTRAVAIL.RetourneLibelle(St_chaine: string;
  bProfil: Boolean): string;
begin
 // bProfil := TRUE > retourne le libelle associé au code Profil
 // sinon retourne le libelle associé au code.
 if bProfil then
   result := RechDom('TTUSERGROUPE',St_chaine,FALSE)
 else
   result := Copy(st_chaine,1,pos('[',st_chaine)-1);
end;
////////////////////////////////////////////////////////////////////////////////
function TOF_GRPTRAVAIL.RetourneProfil(St_chaine: string): string;
var
  StrCh,StrResult : string;
begin
  // Retourne le code profil entre []
  StrResult := '';
  if pos('~',st_chaine) > 0 then
  begin
    StrCh := copy(st_chaine,pos('~',st_chaine)+2,length(st_chaine)-pos('~',st_chaine)+2);
    StrResult := Copy(StrCh,pos('[',StrCh)+1,pos(']',StrCh)-pos('[',StrCh)-1);
  end;
  result := StrResult;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.AddToArrTree (Cod : String ; TreeItm : TTreeNode) ;
begin
   if GetItmFromCode(Cod) = Nil then
   begin
      SetLength(ArrTree,length(ArrTree)+1) ;
      ArrTree[length(ArrTree)-1].Code := Cod ;
      ArrTree[length(ArrTree)-1].TreeItem := TreeItm ;
   end;
end;
////////////////////////////////////////////////////////////////////////////////
function  TOF_GRPTRAVAIL.GetItmFromCode (Cod : String ) : TTreeNode ;
var
  i : integer ;
begin
  Result := nil ;
  for i:=0 to High(ArrTree) do
    if SameText(ArrTree[i].Code,Cod) then
    begin
       Result := ArrTree[i].TreeItem ;
       break ;
    end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.GestionDesDroits;
begin
  // Gestion des Restrictions avec un profil non administrateur.
  if (not v_pgi.Superviseur) then
  begin
    Obj_BtIns.Enabled        := FALSE;
    Obj_BtNIns.Enabled       := FALSE;
    Obj_BtDel.Enabled        := FALSE;
    Obj_BtRen.Enabled        := FALSE;
    Obj_BtSearch.Enabled     := FALSE;
    Obj_BtSearch_aff.Enabled := FALSE;
    Obj_BtAllDnD.Enabled     := FALSE;

    TMenuItem (GetControl ('MENU')).Enabled       := FALSE;
    TMenuItem (GetControl ('SOUS_MENU')).Enabled  := FALSE;
    TMenuItem (GetControl ('SUPPRIMER')).Enabled  := FALSE;
    TMenuItem (GetControl ('RENOMMER')).Enabled   := FALSE;
    TMenuItem (GetControl ('RECHERCHER')).Enabled := FALSE;
    TMenuItem (GetControl ('MULTIAFFEC')).Enabled   := FALSE;

    if (Obj_RBtn_Usr.Checked) {$IFDEF DP}and (JaiLeDroitAdminDossier){$ENDIF} then
    begin
      Obj_BtDnD.Enabled := TRUE;
      TButton   (GetControl ('BVALIDER')).Visible := TRUE;
    end
    else
    begin
      Obj_BtDnD.Enabled := FALSE;
      TButton   (GetControl ('BVALIDER')).Visible := FALSE;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin
 // Se produit quand un noeud va être développé.
 // AllowExpansion = FALSE > empêche le développement du noeud.
 if Node.ImageIndex = Ico_GrDenied then
   AllowExpansion := FALSE;

 Node.Selected := TRUE;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoBeforeExpandNode(Noeud: TTreeNode);
var
 StrCode,StrNode,StrSql,StrNom : string;
 TobAdd,TobPtr                 : Tob;
 NoeudChild                    : TTreeNode;
 IntGrpid                      : integer;
 BooExpand                     : boolean;
                        {<--------------------->}
 procedure RadioBtnDos;
 var
   i : integer;
   TobDos : Tob;
   StrNoDossier,StrDossLib : string;
 begin
    //Creation de la Tob
    TobDos := tob.Create('VirtualTob',nil,-1);
    //Recherche des enfants rattachés au noeud sélectionné.
    if s_ChampLibLiee <> '' then
      StrSql := 'SELECT '+S_tabprfx+'_GRPID,'+S_tabprfx+'_MARK,'+S_tabprfx+'_NODOSSIER,'+
                         S_tabprfx+'_NOM,GRP_CODE "'+S_tabprfx+'_CODE",'+s_ChampLibLiee+' "'+S_tabprfx+'_DOSSLIB" '+
                  'FROM '+S_tablien+
            ' LEFT JOIN GRPDONNEES ON '+S_tablien+'.'+S_tabprfx+'_GRPID = GRPDONNEES.GRP_ID '+
                   'AND '+S_tablien+'.'+S_tabprfx+'_NOM = GRPDONNEES.GRP_NOM '+
      ' LEFT OUTER JOIN '+s_Tabliee+' ON '+S_tablien+'.'+S_tabprfx+'_NODOSSIER = '+s_Tabliee+'.'+s_CleTabLiee+
                ' WHERE '+S_tabprfx+'_NOM = "'+VarIn[0]+'" '+
                   'AND GRP_CODE ="'+StrCode+'" '+
                   'AND '+S_tabprfx+'_MARK ="X"'
    else
      StrSql := 'SELECT '+S_tabprfx+'_GRPID,'+S_tabprfx+'_MARK,'+S_tabprfx+'_NODOSSIER,'+
                         S_tabprfx+'_NOM,GRP_CODE "'+S_tabprfx+'_CODE",null "'+S_tabprfx+'_DOSSLIB" '+
                  'FROM '+S_tablien+
            ' LEFT JOIN GRPDONNEES ON '+S_tablien+'.'+S_tabprfx+'_GRPID = GRPDONNEES.GRP_ID '+
                   'AND '+S_tablien+'.'+S_tabprfx+'_NOM = GRPDONNEES.GRP_NOM '+
                 'WHERE '+S_tabprfx+'_NOM = "'+VarIn[0]+'" '+
                   'AND GRP_CODE ="'+StrCode+'" '+
                   'AND '+S_tabprfx+'_MARK ="X"';

    //Chargement de la Tob
    TobDos.LoadDetailFromSQL(StrSql);

    //Ctrl si existence des données dans Tb_Grpdos.
    TobPtr := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE'],[StrCode],FALSE);
    if TobPtr = nil then
    begin
      Obj_TreeVw.Items.BeginUpdate;
      //Alimentation de la TOB Tb_GrpDos
      for i:=0 to TobDos.Detail.Count-1 do
      begin
        //Init.
        IntGrpid     := TobDos.Detail[i].GetInteger(S_tabprfx+'_GRPID');
        StrNom       := TobDos.Detail[i].GetString(S_tabprfx+'_NOM');
        StrNoDossier := TobDos.Detail[i].GetString(s_CleTabLien);
        StrDossLib   := TobDos.Detail[i].GetString(S_tabprfx+'_DOSSLIB');

        //Avant création, ctrl si l'élément a été supprimé.
        TobPtr := Tb_Erase.FindFirst([S_tabprfx+'_NOM',S_tabprfx+'_GRPID',s_CleTabLien],[StrNom,IntGrpid,StrNoDossier],FALSE);
        if TobPtr = nil then
        begin
          TobAdd := TOB.Create(S_tablien,Tb_GrpDos,-1);
          TobAdd.PutValue(S_tabprfx+'_GRPID',IntGrpid);
          TobAdd.PutValue(s_CleTabLien,StrNoDossier);
          TobAdd.PutValue(S_tabprfx+'_NOM'  ,StrNom);
          TobAdd.PutValue(S_tabprfx+'_MARK' ,'X');
          TobAdd.AddChampSupValeur(S_tabprfx+'_CODE'   ,StrCode);
          TobAdd.AddChampSupValeur(S_tabprfx+'_DOSSLIB',StrDossLib);

          StrNode := StrDossLib+'['+StrNoDossier+']';

          // Ajout du noeud dans la vue.
          NoeudChild := Obj_TreeVw.Items.Addchild(Noeud,StrNode);
          // Init. propriétés du noeud.
          NoeudChild.StateIndex    := 0;
          NoeudChild.ImageIndex    := Ico_Dos;
          NoeudChild.SelectedIndex := Ico_Dos;
          BooExpand := TRUE;
        end;
        Application.ProcessMessages;
      end;
      Obj_TreeVw.Items.EndUpdate;
    end;
    // Destruction
    FreeAndNil(TobDos);
 end;
                        {<--------------------->}
 procedure RadioBtnUsr;
 var
   i : integer;
   TobUsr : TOB;
   StrProfil,StrUser,StrUserLib,StrClause : String;
 begin
    //creation de la TOB
    TobUsr := tob.Create('VirtualTob',nil,-1);

    //si l'utilisateur en cours n'est pas admin l'exclure des groupes de travail.
    if not v_pgi.Superviseur then
      StrClause := ' AND LND_USERID <> "'+v_pgi.User+'"'
    else
      StrClause := '';

    StrSql := 'SELECT LND_GRPID,LND_PROFILID,LND_USERID,'+
                     'LND_NOM,GRP_CODE "LND_CODE",US_LIBELLE "LND_USERLIB" '+
                'FROM LIENDONNEES '+
           'LEFT JOIN GRPDONNEES ON LIENDONNEES.LND_GRPID = GRPDONNEES.GRP_ID '+
                 'AND LIENDONNEES.LND_NOM = GRPDONNEES.GRP_NOM '+
	   'LEFT OUTER JOIN UTILISAT ON LIENDONNEES.LND_USERID = UTILISAT.US_UTILISATEUR '+
               'WHERE LND_NOM = "'+VarIn[0]+'" '+
                 'AND GRP_CODE ="'+StrCode+'"'+
                 StrClause;
    //Chargement de la TOB.
    TobUsr.LoadDetailFromSQL(StrSql);

    //Ctrl si existence des données dans Tb_TobUsr.
    TobPtr := Tb_GrpUsr.FindFirst(['LND_CODE'],[StrCode],FALSE);
    if TobPtr = nil then
    begin
      Obj_TreeVw.Items.BeginUpdate;
      for i:=0 to TobUsr.Detail.Count-1 do
      begin
        //Alimentation de la TOB Tb_GrpUsr
        IntGrpid  := TobUsr.Detail[i].GetInteger('LND_GRPID');
        StrProfil := TobUsr.Detail[i].GetString('LND_PROFILID');
        StrUser   := TobUsr.Detail[i].GetString('LND_USERID');
        StrNom    := TobUsr.Detail[i].GetString('LND_NOM');
        StrCode   := TobUsr.Detail[i].GetString('LND_CODE');
        StrUserLib:= TobUsr.Detail[i].GetString('LND_USERLIB');
        //Avant création, ctrl si l'élément a été supprimé.
        TobPtr := Tb_Erase.FindFirst(['LND_NOM','LND_GRPID','LND_PROFILID','LND_USERID'],[StrNom,IntGrpid,StrProfil,StrUser],FALSE);
        if TobPtr = nil then
        begin
          TobAdd := TOB.Create('LIENDONNEES',Tb_GrpUsr,-1);
          TobAdd.PutValue('LND_GRPID'   ,IntGrpid);
          TobAdd.PutValue('LND_PROFILID',StrProfil);
          TobAdd.PutValue('LND_USERID'  ,StrUser);
          TobAdd.PutValue('LND_NOM'     ,StrNom);
          TobAdd.AddChampSupValeur('LND_CODE',StrCode);

          //Recherche du profil
          if StrProfil <> '' then
            StrNode  := StrUserLib+'['+StrUser+'] ~ '+RetourneLibelle(StrProfil,true)+'['+StrProfil+']'
          else
            StrNode := StrUserLib+'['+StrUser+']';

          NoeudChild := Obj_TreeVw.Items.Addchild(Noeud,StrNode);
          // Init. propriétés du noeud.
          NoeudChild.StateIndex    := 1;
          NoeudChild.ImageIndex    := Ico_Usr;
          NoeudChild.SelectedIndex := Ico_Usr;
          BooExpand := TRUE;
        end;
        Application.ProcessMessages;
      end;
      Obj_TreeVw.Items.EndUpdate;
    end;
    FreeAndNil(TobUsr);
 end;
                        {<--------------------->}
 begin
  // Se produit quand un noeud a été développé.
  BooExpand := FALSE;
  StrCode  := RetourneCode(Noeud.Text);

  //exécution du traitement en fonction du choix du groupe de radio bouton
  if Obj_RBtn_Dos.Checked then
    RadioBtnDos
  else if Obj_RBtn_Usr.Checked then
    RadioBtnUsr
  else if Obj_RBtn_All.Checked then
  begin
    RadioBtnDos;
    RadioBtnUsr;
  end;
  if BooExpand then
    Noeud.Expand(false);
end;
////////////////////////////////////////////////////////////////////////////////
function TOF_GRPTRAVAIL.ExisteAffectation(St_Code,St_Donnee: string): Boolean;
var
  //StrSql,StrCls_1,StrCls_2 : string;
  tb_tmp : TOB;
  BooExisteSQL : boolean;
begin
  Result := FALSE;
  BooExisteSQL := FALSE;
 {StrCls_1 := '';
  StrCls_2 := '';}

 {if St_Donnee <> '' then
  begin
    StrCls_1 := 'AND LND_USERID ="'+St_Donnee+'" ';
    StrCls_2 := 'AND '+s_CleTabLien+'="'+St_Donnee+'" ';
  end;}

  // Ctrl LienDosGrp
  if St_Donnee <> '' then
    tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE',S_tabprfx+'_NODOSSIER'],[St_Code,St_Donnee],FALSE)
  else
    tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE'],[St_Code],FALSE);

  if tb_tmp <> nil then
    BooExisteSQL := TRUE;

  // Ctrl LienDonnees
  if St_Donnee <> '' then
    tb_tmp := Tb_GrpUsr.FindFirst(['LND_CODE','LND_USERID'],[St_Code,St_Donnee],FALSE)
  else
    tb_tmp := Tb_GrpUsr.FindFirst(['LND_CODE'],[St_Code],FALSE);

  if tb_tmp <> nil then
    BooExisteSQL := TRUE;

  //Ctrl si existence d'affectation avant suppression d'un groupe de travail.
  {StrSql := 'SELECT 1 '+
              'FROM LIENDONNEES '+
         'LEFT JOIN GRPDONNEES ON LIENDONNEES.LND_GRPID = GRPDONNEES.GRP_ID '+
                             'AND LIENDONNEES.LND_NOM   = GRPDONNEES.GRP_NOM '+
             'WHERE LND_NOM = "'+VarIn[0]+'" '+
               'AND GRP_CODE ="'+St_Code+'" '+
               StrCls_1+
             'UNION '+
            'SELECT 1 '+
              'FROM '+S_tablien+
        ' LEFT JOIN GRPDONNEES ON '+S_tablien+'.'+S_tabprfx+'_GRPID = GRPDONNEES.GRP_ID '+
                             'AND '+S_tablien+'.'+S_tabprfx+'_NOM   = GRPDONNEES.GRP_NOM '+
             'WHERE '+S_tabprfx+'_NOM = "'+VarIn[0]+'" '+
               'AND GRP_CODE ="'+St_Code+'" '+
               'AND '+S_tabprfx+'_MARK ="X" '+
               StrCls_2;}

  // Retourne TRUE si la requête SQL possède des enregistrements.
  if BooExisteSQL then
    Result := TRUE;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.RechercheNoeud(sender : TObject;St_donnee: string);
var
  TobDos,TobUsr : TOB;
  StrSql,StrCode : string;
  Noeud : TTreeNode;
  idx : Integer;
begin
  if Obj_RBtn_Dos.Checked then
  begin
    TobDos := TOB.Create('VirtualTob',nil,-1);
    StrSql := 'SELECT GRP_CODE '+
                'FROM '+S_tablien+
          ' LEFT JOIN GRPDONNEES ON '+S_tablien+'.'+S_tabprfx+'_GRPID = GRPDONNEES.GRP_ID '+
                 'AND '+S_tablien+'.'+S_tabprfx+'_NOM = GRPDONNEES.GRP_NOM '+
               'WHERE '+S_tabprfx+'_NOM = "'+VarIn[0]+'" '+
                 'AND '+S_tabprfx+'_NODOSSIER = "'+St_donnee+'" '+
                 'AND '+S_tabprfx+'_MARK ="X"';
    TobDos.LoadDetailFromSQL(StrSql);

    for idx:=0 to TobDos.Detail.Count-1 do
    begin
      StrCode := TobDos.Detail[idx].GetString('GRP_CODE');
      Noeud := GetItmFromCode(StrCode);
      DoBeforeExpandNode(Noeud);
    end;

    FreeAndNil(TobDos);
  end
  else if Obj_RBtn_Usr.Checked then
  begin
    TobUsr := TOB.Create('VirtualTob',nil,-1);

    StrSql := 'SELECT GRP_CODE '+
                'FROM LIENDONNEES '+
           'LEFT JOIN GRPDONNEES ON LIENDONNEES.LND_GRPID = GRPDONNEES.GRP_ID '+
                 'AND LIENDONNEES.LND_NOM = GRPDONNEES.GRP_NOM '+
               'WHERE LND_NOM = "'+VarIn[0]+'" '+
                 'AND LND_USERID ="'+St_donnee+'"';
    //Chargement de la TOB.
    TobUsr.LoadDetailFromSQL(StrSql);

    for idx:=0 to TobUsr.Detail.Count-1 do
    begin
      StrCode := TobUsr.Detail[idx].GetString('GRP_CODE');
      Noeud := GetItmFromCode(StrCode);
      DoBeforeExpandNode(Noeud);
    end;

    FreeAndNil(TobUsr);
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.AddToArrSearch(ItemList: String;
  IdColor: integer);
begin
   if GetIndexFromItem(ItemList) = -1 then
   begin
      SetLength(ArrSearch,length(ArrSearch)+1);
      ArrSearch[length(ArrSearch)-1].StrItemList := ItemList;
      ArrSearch[length(ArrSearch)-1].IdxColor    := IdColor;
   end;
end;
////////////////////////////////////////////////////////////////////////////////
function TOF_GRPTRAVAIL.GetIndexFromItem(ItemList: String): Integer;
var
  i : integer ;
begin
  Result := -1 ;
  for i:=0 to High(ArrSearch) do
    if SameText(ArrSearch[i].StrItemList,ItemList) then
    begin
      Result := ArrSearch[i].IdxColor;
      break;
    end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoBeforeDragnDrop;
var
  Fecran : TForm;
  idx    : integer;
  Tb_tmp : TOB;
  s_ColValue,s_ColLibValue,s_Value,s_LibValue : string;
  b_ColDisplay : boolean;

  {$IFDEF EAGLCLIENT}
      HGrid : THGrid;
  {$ELSE}
      HGrid : THDBGrid;
  {$ENDIF}
begin
  Fecran := obj_Pnlst.InsideForm;
  HGrid  := TFMul(Fecran).FListe;
  b_ColDisplay := FALSE;
  Tb_MultiSelect.ClearDetail;

  // Initialisation des noms de colonnes
  if Obj_RBtn_Dos.Checked then
  begin
    s_ColValue    := s_CleTabLiee;
    s_ColLibValue := s_ChampLibLiee;
  end
  else if Obj_RBtn_Usr.Checked then
  begin
    s_ColValue    := 'US_UTILISATEUR';
    s_ColLibValue := 'US_LIBELLE';
  end;

  //vérification si champ "s_ColLibValue" visible dans la liste.
  {$IFDEF EAGLCLIENT}
    for idx:=0 to HGrid.ColCount -1 do
    begin
      if SameText(s_ColLibValue,HGrid.FColNames[idx]) then
      begin
        b_ColDisplay := TRUE;
        Break;
      end;
    end;
  {$ELSE}
    for idx:=0 to HGrid.FieldCount-1 do
    begin
      if SameText(s_ColLibValue,HGrid.Columns.Items[idx].FieldName) then
      begin
        b_ColDisplay := TRUE;
        Break;
      end;
    end;
  {$ENDIF}

  if HGrid.nbSelected > 1 then
  begin
    //Modif l'apparence sur curseur.
    HGrid.DragCursor := crMultiDrag;

    // Récupére la sélection de la liste
    for idx :=0 to HGrid.nbSelected -1 do
    begin
      HGrid.GotoLeBookmark(idx);

      {$IFDEF EAGLCLIENT}
        TFMul(Fecran).Q.TQ.Seek(HGrid.Row-1);
      {$ENDIF}

      s_Value := TFMul(Fecran).Q.FindField(s_ColValue).AsString ;
      if b_ColDisplay then
        s_LibValue := TFMul(Fecran).Q.FindField(s_ColLibValue).AsString
      else
        s_LibValue := '';

      Tb_tmp := tob.create('',Tb_MultiSelect,-1);
      Tb_tmp.AddChampSupValeur('MUL_SELECT',s_Value);
      Tb_tmp.AddChampSupValeur('MUL_LIBELLE',s_LibValue);
      Tb_tmp.AddChampSupValeur('MUL_AUTORISE','X');
    end;
  end
  else if HGrid.AllSelected then
  begin
    HGrid.DragCursor := crMultiDrag;

    TFMul(Fecran).Q.First;
    while Not TFMul(Fecran).Q.EOF do
    begin
      s_Value := TFMul(Fecran).Q.FindField(s_ColValue).AsString ;
      if b_ColDisplay then
        s_LibValue := TFMul(Fecran).Q.FindField(s_ColLibValue).AsString
      else
        s_LibValue := '';

      Tb_tmp := tob.create('',Tb_MultiSelect,-1);
      Tb_tmp.AddChampSupValeur('MUL_SELECT',s_Value);
      Tb_tmp.AddChampSupValeur('MUL_LIBELLE',s_LibValue);
      Tb_tmp.AddChampSupValeur('MUL_AUTORISE','X');

      Application.ProcessMessages;
      TFMul(Fecran).Q.Next;
    end;
  end
  else
  begin
    HGrid.DragCursor := crDrag;

    {$IFDEF EAGLCLIENT}
      TFMul(Fecran).Q.TQ.Seek(HGrid.Row-1);
    {$ENDIF}

    s_Value := TFMul(Fecran).Q.FindField(s_ColValue).AsString ;
    if b_ColDisplay then
      s_LibValue := TFMul(Fecran).Q.FindField(s_ColLibValue).AsString
    else
      s_LibValue := '';

    Tb_tmp := tob.create('',Tb_MultiSelect,-1);
    Tb_tmp.AddChampSupValeur('MUL_SELECT',s_Value);
    Tb_tmp.AddChampSupValeur('MUL_LIBELLE',s_LibValue);
    Tb_tmp.AddChampSupValeur('MUL_AUTORISE','X');
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
 {$IFDEF EAGLCLIENT}
  HGrid : THGrid;
 {$ELSE}
  HGrid : THDBGrid;
 {$ENDIF}
begin
  HGrid := nil;

  {$IFDEF EAGLCLIENT}
  if (sender is THGrid) then
    HGrid := (sender as THGrid);
  {$ELSE}
  if (sender is THDBGrid) then
    HGrid := (sender as THDBGrid);
  {$ENDIF}

  // Si bouton gauche de la souris maintenue enfoncé + déplacement de la souris =
  // activation du drag n drop SAUF dans le cas d'un Resize colonne.
  if (ssLeft in	Shift) then
  begin
    // activation DragnDrop sur la liste
    if HGrid <> nil then
    begin
      if (v_pgi.Superviseur) or ((not v_pgi.Superviseur) and (Obj_RBtn_Usr.Checked) {$IFDEF DP} and (JaiLeDroitAdminDossier){$ENDIF}) then
      begin
        if HGrid.MouseCoord(X,Y).Y > 0 then
          HGrid.BeginDrag(false);
      end;
    end
    //activation DragnDrop sur l'arborescence.
    else if (sender is TTreeView) then
    begin
      if v_pgi.Superviseur then
        Obj_TreeVw.BeginDrag(false);
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.AddToArrNode(var ArrayNode:TArrayOfTreeNode ; Noeud : TTreeNode);
  function ExisteNode(Noeud : TTreeNode) : Boolean;
  var
    i : integer;
  begin
    Result := FALSE ;
    for i:=0 to High(ArrayNode) do
      if ArrayNode[i].ItemId = Noeud.ItemId then
      begin
        Result := TRUE;
        break ;
      end;
  end;
begin
  //Ctrl si l'élément déjà inséré dans le tableau
  if not ExisteNode(Noeud) then
  begin
    SetLength(ArrayNode,length(ArrayNode)+1);
    ArrayNode[length(ArrayNode)-1] := Noeud;
//    SetLength(ArrNode,length(ArrNode)+1);
//    ArrayNode[length(ArrayNode)-1] := Noeud;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoDeleteArrNode(Sender: TObject);
var
  i : integer;
  BooDel : boolean;
begin
  if PGIAsk('Êtes-vous certain de vouloir retirer le ou les élément(s) sélectionné(s) ?',Ecran.Caption)=mrYes then
  begin
    BooDel := FALSE;
    for i:=0 to High(ArrNode) do
    begin
      if ArrNode[i].ImageIndex = ico_delete then
      begin
        case ArrNode[i].StateIndex of
          0 : SuppAffectationDos(ArrNode[i]);
          1 : SuppAffectationUsr(ArrNode[i]);
        end;
        BooDel := TRUE;
        ArrNode[i].Delete;
      end;
    end;
    SetLength(ArrNode,0);
    if not BooDel then
      PGIInfo('Aucun élément n''a été retiré de l''arborescence.'+#13#10+
              'Veuillez sélectionner le ou les éléments à l''aide de la barre d''espace.')
    else
      B_flg_Valid := FALSE;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.SuppAffectationDos(Noeud : TTreeNode);
var
  NoeudParent,NoeudTemp : TTreeNode;
  GrpId,idx : Integer;
  Tb_tmp,Tb_del,Tb_comp,Tb_temp : TOB;
  S_enfant,s_code,S_Sql : String;
  b_DelRight : boolean;
begin
  // suppression d'une affectation DOSSIER.
  // Avant suppression, vérifier si le dossier est affecté à un autre groupe de travail.
  // si NON -> suppression de l'affectation marquée 'X' + affectations liées aux noeuds parent (marquées '-').
  // si OUI -> récuperer l'arborescence du ou des dossiers affectés aux autres groupes de travail et sauvegarder les correspondances.
  GrpId    := 0;
  Tb_comp  := TOB.Create('VirtualTob',nil,-1);
  s_code   := RetourneCode(Noeud.Parent.Text);
  S_enfant := RetourneCode(Noeud.Text);
  b_DelRight := TRUE; //autorise la suppression.

  S_Sql  := 'SELECT GRP_CODE '+
              'FROM '+S_tablien+
        ' LEFT JOIN GRPDONNEES ON '+S_tablien+'.'+S_tabprfx+'_GRPID = GRPDONNEES.GRP_ID '+
               'AND '+S_tablien+'.'+S_tabprfx+'_NOM = GRPDONNEES.GRP_NOM '+
             'WHERE '+S_tabprfx+'_NOM = "'+VarIn[0]+'" '+
               'AND GRP_CODE <> "'+s_code+'" '+
               'AND '+S_tabprfx+'_MARK ="X" '+
               'AND '+S_tabprfx+'_NODOSSIER ="'+S_enfant+'"';

  if ExisteSQL(S_Sql) then
  begin
    Tb_comp.LoadDetailFromSQL(S_sql);
    for idx := 0 to Tb_comp.Detail.Count-1 do
    begin
      s_code := Tb_comp.Detail[idx].GetString('GRP_CODE');
      NoeudTemp := GetItmFromCode(s_code);
      NoeudParent := NoeudTemp.Parent;
      while NoeudParent <> nil do
      begin
        s_code := RetourneCode(NoeudParent.Text);
        Tb_tmp := TOB.Create('',Tb_comp,-1);
        Tb_tmp.AddChampSupValeur('GRP_CODE',s_code);
        NoeudParent := NoeudParent.Parent;
      end;
      Application.ProcessMessages;
    end;
  end;

  //Suppression de l'affectation en cours 'X' + affectations rattachées aux différents parents '-'.
  NoeudParent := Noeud.Parent;
  while (NoeudParent <> nil)  do
  begin
    s_code := RetourneCode(NoeudParent.Text);

    //verif si existence du noeud parent en correspondance avec la Tb_comp.
    if Tb_comp.Detail.Count > 0 then
    begin
      Tb_tmp := Tb_comp.FindFirst(['GRP_CODE'],[s_code],FALSE);
      if Tb_tmp <> nil then
        b_DelRight := FALSE;
    end;

    if b_DelRight then
    begin
      Tb_tmp := Tb_GrpDos.FindFirst([S_tabprfx+'_CODE',s_CleTabLien],[s_code, S_enfant],FALSE);
      if Tb_tmp <> nil then
      begin
        Tb_del := tob.create(S_tablien,tb_erase,-1);
        Tb_del.PutValue(S_tabprfx+'_GRPID',Tb_tmp.GetInteger(S_tabprfx+'_GRPID'));
        Tb_del.PutValue(S_tabprfx+'_MARK',Tb_tmp.GetString(S_tabprfx+'_MARK'));
        Tb_del.PutValue(s_CleTabLien,Tb_tmp.GetString(s_CleTabLien));
        Tb_del.PutValue(S_tabprfx+'_NOM',VarIn[0]);
        Tb_tmp.free;
        //retrait dans la TOB TB_InsOrUpd
        Tb_temp := Tb_InsOrUpd.FindFirst([S_tabprfx+'_CODE',s_CleTabLien],[s_code, S_enfant],FALSE);
        if tb_temp <> nil then
          tb_temp.free;
      end
      else
      begin
        // suppression des affectations marquées '-'
        Tb_tmp := Tb_GrpTrav.FindFirst(['GRP_CODE'],[s_code],FALSE);
        if Tb_tmp <> nil then
          GrpId := Tb_tmp.GetInteger('GRP_ID');
        Tb_del := tob.create(S_tablien,tb_erase,-1);
        Tb_del.PutValue(S_tabprfx+'_GRPID',GrpId);
        Tb_del.PutValue(S_tabprfx+'_MARK','-');
        Tb_del.PutValue(s_CleTabLien,S_enfant);
        Tb_del.PutValue(S_tabprfx+'_NOM',VarIn[0]);
      end;
    end;
    NoeudParent := NoeudParent.parent;
    Application.ProcessMessages;
  end;

  FreeAndNil(Tb_comp);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.SuppAffectationUsr(Noeud : TTreeNode);
var
  NoeudParent,NoeudEnfant,NoeudSuivant,NoeudTemp : TTreeNode;
  Tb_tmp,Tb_del,tb_temp : TOB;
  s_noeud,S_enfant,s_code,S_Profil : String;
begin
  //Suppression de l'utilisateur ainsi que ses occurrences(user/profil)
  //affectées aux noeuds enfant.
  s_noeud      := Noeud.Text;
  S_enfant     := RetourneCode(Noeud.Text);
  NoeudParent  := Noeud.Parent;
  NoeudSuivant := NoeudParent.getNextSibling;
  NoeudEnfant  := NoeudParent.getFirstChild;

  s_profil := RetourneProfil(Noeud.Text);

  while (NoeudEnfant <> NoeudSuivant) and (NoeudParent.level < NoeudEnfant.Level) do
  begin
    if (NoeudEnfant.StateIndex = 1) and (SameText(NoeudEnfant.text,s_noeud)) then
    begin
      //maj de la tob
      s_code := RetourneCode(NoeudEnfant.Parent.Text);
      Tb_tmp := Tb_GrpUsr.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,S_enfant,s_profil],FALSE);
      if Tb_tmp <> nil then
      begin
        Tb_del := tob.create('LIENDONNEES',tb_erase,-1);
        Tb_del.PutValue('LND_GRPID',Tb_tmp.GetInteger('LND_GRPID'));
        Tb_del.PutValue('LND_PROFILID',Tb_tmp.GetString('LND_PROFILID'));
        Tb_del.PutValue('LND_USERID',Tb_tmp.GetString('LND_USERID'));
        Tb_del.PutValue('LND_NOM',VarIn[0]);
        Tb_tmp.free;
        //
        tb_temp := Tb_InsOrUpd.FindFirst(['LND_CODE','LND_USERID','LND_PROFILID'],[s_code,S_enfant,s_profil],FALSE);
        if tb_temp <> nil then
          tb_temp.free;
      end;
      //suppression du noeud dans la vue.
      NoeudTemp   := NoeudEnfant;
      NoeudEnfant := NoeudEnfant.GetNext;
      NoeudTemp.Delete;
    end
    else
    begin
      if NoeudEnfant.StateIndex = -1 then
        DoBeforeExpandNode(NoeudEnfant);
      NoeudEnfant := NoeudEnfant.GetNext;
    end;
    Application.ProcessMessages;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickBInsert(Sender: TObject);
var
  f_form : TForm;
begin
  f_form := obj_Pnlst.InsideForm;
  AglLanceFiche('YY','YYUTILISAT','','','ACTION=CREATION');
  AGLRefreshDB([LongInt(f_form), 'FListe'], 2);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.RechercheNoeudEnfants(NoeudParent,
  Noeudenfant: TTreeNode);
var i : integer;
    NoeudEnf : TTreeNode;
begin
  for i:=0 to NoeudParent.Count-1 do
  begin
    NoeudEnf := Obj_TreeVw.Items.AddChild(Noeudenfant,NoeudParent.Item[i].Text);
    NoeudEnf.StateIndex    := NoeudParent.Item[i].StateIndex;
    NoeudEnf.SelectedIndex := NoeudParent.Item[i].SelectedIndex;
    NoeudEnf.ImageIndex    := NoeudParent.Item[i].ImageIndex;

    if NoeudParent.Item[i].Count > 0 then
      RechercheNoeudEnfants(NoeudParent.Item[i],NoeudEnf);
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickBvoir(Sender: TObject);
var
  f_form : TForm;
  StrUser : string;
  {$IFDEF EAGLCLIENT}
    HGrid : THGrid;
  {$ELSE}
    HGrid : THDBGrid;
  {$ENDIF}
begin
  //Ouvre la fiche signalétique de l'utilisateur sélectionné dans la liste.
  f_form := obj_Pnlst.InsideForm;
  HGrid  := TFMul(f_form).FListe;

  {$IFDEF EAGLCLIENT}
    TFMul(f_form).Q.TQ.Seek(HGrid.Row-1);
  {$ENDIF}

  StrUser := TFMul(f_form).Q.FindField('US_UTILISATEUR').AsString;

  AglLanceFiche('YY','YYUTILISAT','',StrUser,'ACTION=MODIFICATION');
  AGLRefreshDB([LongInt(f_form), 'FListe'], 2);
end;
////////////////////////////////////////////////////////////////////////////////
procedure AffecteDosGrpConf (sNoDossier: string; bDelete : Boolean = FALSE);
//Affecte le groupe de travail au dossier pour qu'il apparaisse dans la liste des dossiers
var
  Q:TQuery;
  St:string;
begin
  if sNoDossier = '' then exit;

  Q := OpenSQL('SELECT GRP_CODE AS DOG_GROUPECONF, LDO_NODOSSIER AS DOG_NODOSSIER FROM GRPDONNEES, LIENDOSGRP'
            +  ' WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID'
            +  ' AND LDO_MARK = "X" AND LDO_NODOSSIER = "' + SNoDossier + '"',True);

  //Si aucun enregistrement trouvé ou erreur, on force DOS_GROUPECONF à vide
  if Q.RecordCount < 1 then ExecuteSql('UPDATE DOSSIER SET DOS_GROUPECONF = ""'
                                    + ' WHERE DOS_NODOSSIER = "' + sNoDossier + '"');

  //Si plusieurs enregistrements trouvés, on force DOS_GROUPECONF à "X" pour indiquer le multi-groupes
  if Q.RecordCount > 1 then ExecuteSql('UPDATE DOSSIER SET DOS_GROUPECONF = "X"'
                                    + ' WHERE DOS_NODOSSIER = "' + sNoDossier + '"');

  //Si un seul enregistrement trouvé, on affecte le groupe normalement
  if Q.RecordCount = 1 then
  begin
    if not bDelete then
    begin
      St := Q.FindField('DOG_GROUPECONF').AsString;
      ExecuteSql('UPDATE DOSSIER SET DOS_GROUPECONF = "' + St + '"'
              + ' WHERE DOS_NODOSSIER = "' + sNoDossier + '"');
    end
    else
      ExecuteSql('UPDATE DOSSIER SET DOS_GROUPECONF = ""'
              + ' WHERE DOS_NODOSSIER = "' + sNoDossier + '"');
  end;

  Ferme(Q);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.LanceFenetreRecherche(ItemState : integer);
var
  LstCol  : TListColumn;
  LstItm  : TListItem;
  Noeud   : TTreeNode;
  StrLib,StrGrpT,s_Search,StrProf,StrLibProf : string;
  idx     : integer;
  booFind : boolean;
begin
  booFind := FALSE;
  if assigned(WinFindResult) then
  begin
    FreeAndNil(LstView);
    FreeAndNil(WinFindResult);
  end;

  // Affiche une fenêtre de type Bloc Note avec les infos de recherche.
  WinFindResult := TToolWindow97.Create(nil);
  with WinFindResult do
  begin
    Parent  := Ecran;
    Caption := 'Résultats de la recherche';
    Left    := 132;
    Top     := 344;
    Height  := 172;
    width   := 500;
  end;

  LstView := TListView.Create(nil);
  with LstView do
  begin
    Parent           := WinFindResult;
    Align            := AlClient;
    ViewStyle        := vsReport;
    RowSelect        := TRUE;
    ReadOnly         := TRUE;
    SortType         := stText;
    OrderSort        := [Croissant];
    SmallImages      := V_PGI.GraphList;
    OnClick          := DoOnClick;
    OnCustomDrawItem := DoOnCustomDrawItem;
    OnColumnClick    := DoOnColumnClick;
    OnCompare        := DoOnCompare;

    //Mise en place des colonnes
    case ItemState of
      0 : begin
            LstCol := Columns.Add;
            LstCol.Caption := 'Dossier';
            LstCol.Width  := 90;

            LstCol := Columns.Add;
            LstCol.Caption := 'Libellé';
            LstCol.Width  := 200;

            LstCol := Columns.Add;
            LstCol.Caption := 'Groupe de travail';
            LstCol.Width  := 200;
          end;
      1 : begin
            LstCol := Columns.Add;
            LstCol.Caption := 'Code utilisateur';
            LstCol.Width  := 90;

            LstCol := Columns.Add;
            LstCol.Caption := 'Nom';
            LstCol.Width  := 150;

            LstCol := Columns.Add;
            LstCol.Caption := 'Profil';
            LstCol.Width  := 50;

            LstCol := Columns.Add;
            LstCol.Caption := 'Groupe de travail';
            LstCol.Width  := 200;
          end;
    end;

    // Affichage de la recherche
    for idx := 0 to Obj_TreeVw.Items.Count -1 do
    begin
      Noeud := Obj_TreeVw.Items[idx];
      if Noeud.StateIndex <> -1 then
      begin
        s_Search := RetourneCode(Noeud.Text);
        if GetIndexFromItem(s_Search) <> -1 then
        begin
            booFind := true;
            if not Obj_TreeVw.Focused then
              Obj_TreeVw.SetFocus;
            Noeud.Selected := TRUE;

            StrLib  := RetourneLibelle(Noeud.Text);
            StrGrpT := Noeud.Parent.text;

            LstItm := Items.Add;
            LstItm.ImageIndex := Ico_PtVert;
            LstItm.Caption := s_Search;   //Col Dossier ou code utilisateur
            LstItm.SubItems.Add(StrLib);  //Col Libelle dossier ou utilisateur
            //cas utilisateur
            if ItemState = 1 then
            begin
              StrProf    := RetourneProfil(Noeud.Text);
              StrLibProf := RetourneLibelle(StrProf,TRUE);
              LstItm.SubItems.Add(StrLibProf); //Col Profil
            end;
            LstItm.SubItems.Add(StrGrpT); //Col Groupe de travail
            LstItm.Data := Noeud;
        end;
        Application.ProcessMessages;
      end;
    end;

    if not booFind then
    begin
      LstView.ViewStyle := vsSmallIcon;
      LstItm := Items.Add;
      LstItm.ImageIndex := 1;       //icone Croix rouge
      LstItm.StateIndex := 1;
      LstItm.Caption := 'La recherche est terminée. Il n''y a aucun résultat à afficher.'
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoOnClick(Sender: TObject);
var
  Item: TListItem;
  Noeud : TTreeNode;
begin
  //Evénement lors du clique sur un élément de la listview.
  if (sender is TListView) then
  begin
    Item := (Sender as TListView).Selected;
    if Item <> nil then
    begin
      Noeud := Item.Data;
      if Noeud <> nil then
      begin
        if not Obj_TreeVw.Focused then
          Obj_TreeVw.SetFocus;

        if not assigned(Obj_TreeVw.Items.GetNode(Noeud.ItemId)) then
        begin
          PGIINFO('Cette information n''est plus disponible dans l''arborescence. Veuillez effectuer une recherche pour rafraichir la liste.',ecran.Caption);
          Item.ImageIndex := Ico_PtRouge;
        end
        else
          Noeud.Selected := TRUE;

        Obj_TreeVw.Refresh;
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoOnCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  //Modifie l'apparence sur la selection d'un item sur la Listview.
  if (Item.Selected) and (item.StateIndex = -1) then
  begin
    Sender.Canvas.Font.Style := [fsBold];
    Sender.Canvas.Font.Color := clBlue;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoOnColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  //Se produit quand l'utilisateur clique sur un en-tête de colonne de la vue liste.
  if croissant in OrderSort then
    OrderSort := [Decroissant]
  else
    OrderSort := [Croissant];

  ColumnToSort := Column.Index;
  (Sender as TListView).AlphaSort;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoOnCompare(Sender: TObject; Item1,
  Item2: TListItem; Data: Integer; var Compare: Integer);
var
  idcol: Integer;
begin
  //Se produit quand deux éléments de la liste doivent être comparés durant
  //le tri de la liste.
  if ColumnToSort = 0 then
  begin
    if Croissant in OrderSort then
      Compare := CompareText(Item1.Caption,Item2.Caption)
    else
      Compare := -CompareText(Item1.Caption,Item2.Caption);
  end
  else
  begin
    idcol := ColumnToSort - 1;
    if Croissant in OrderSort then
      Compare := CompareText(Item1.SubItems[idcol],Item2.SubItems[idcol])
    else
      Compare := -CompareText(Item1.SubItems[idcol],Item2.SubItems[idcol]);
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoActiveMulAff(Sender: TObject);
var
 key : word;
 Shift: TShiftState;
begin
  key := Ord('X');
  Shift := [ssCtrl,ssAlt];
  DoKeyDown(Obj_TreeVw,key,Shift);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoAnnuleMulAff(Sender: TObject);
var
 key : word;
 Shift: TShiftState;
begin
  key := Ord('A');
  Shift := [ssCtrl,ssAlt];
  DoKeyDown(Obj_TreeVw,key,Shift);
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_GRPTRAVAIL.DoClickAllDragnDrop(Sender: TObject);
var
  i,state : integer;
  t,DupGrpT : tob;
  sCode,sAsk : String;
  procedure ExecuteAffectation(Noeud : TTreeNode);
  begin
    if noeud <> nil then
    begin
      //Autorisation des affectations
      AutoriseAffectation(noeud);
      //Ajout des items enfant Sans profil pour les utilisateurs
      AjoutEnfant(noeud,FALSE);
      //Mise des tobs
      MiseAJour_TOB(noeud);
    end;
  end;

begin
  //récupére le ou les élements de la liste.
  DoBeforeDragnDrop;

  //Avertissement si aucun élément sélectionné.
  if Tb_MultiSelect.Detail.Count = 0 then
  begin
    PGIInfo('Veuillez sélectionner un ou plusieurs éléments de la liste',ecran.Caption);
    exit;
  end;

  if Obj_RBtn_Dos.Checked then
  begin
    sAsk := 'Êtes-vous sûr de vouloir affecter la sélection à tous les groupes de travail ?';
    state := 0;
  end
  else if Obj_RBtn_Usr.Checked then
  begin
    sAsk := 'Êtes-vous sûr de vouloir affecter la sélection à tous les groupes de travail ?'+#13#10+
            'Avertissement ! Aucun profil attribué sur l''utilisateur.';
    state := 1;
  end;

  // Affectation à tous les groupes de travail.
  if PGIAsk(sAsk,ecran.caption)=mrYes then
  begin
    InitMoveProgressForm(nil,ecran.Caption,'Affectation en cours, veuillez patienter...',1,FALSE,TRUE);
    Obj_TreeVw.Items.BeginUpdate;

    case State of
      // Affectation DOSSIER
      0 : begin
            MaxValueProgressForm(High(ArrTree));

            for i:=0 to High(ArrTree) do
            begin
              ExecuteAffectation(ArrTree[i].TreeItem);
              MoveCurProgressForm;
              Application.ProcessMessages;
            end;
          end;
      // Affectation Utilisateur.
      1 : begin
            // on duplique la tob Tb_GrpTrav car dans les traitements
            // qui suivent, on effectue aussi une recherche sur Tb_GrpTrav.
            DupGrpT := Tob.Create('DuplicateTob',nil,-1);
            DupGrpT.Dupliquer(Tb_GrpTrav,TRUE,TRUE,TRUE);
            i := round(DupGrpT.SommeSimple('GRP_NOM',['GRP_IDPERE'],[-1],TRUE));
            MaxValueProgressForm(i);

            // Recherche des items dans la vue Hier. qui n'ont pas de parents.
            t := DupGrpT.FindFirst(['GRP_IDPERE'],[-1],FALSE);
            while t <> nil do
            begin
              sCode := t.GetString('GRP_CODE');
              ExecuteAffectation(GetItmFromCode(sCode));
              MoveCurProgressForm;
              Application.ProcessMessages;
              t := DupGrpT.FindNext(['GRP_IDPERE'],[-1],FALSE);
            end;
            FreeAndNil(DupGrpT);
          end;
    end;
    FiniMoveProgressForm;
    Obj_TreeVw.Items.EndUpdate;
  end;

end;

Initialization
  registerclasses ( [ TOF_GRPTRAVAIL ] ) ;
end.
