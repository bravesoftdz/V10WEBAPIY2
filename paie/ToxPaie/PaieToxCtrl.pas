{***********UNITE*************************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Gestion des corbeilles de la TOX
Mots clefs ... : CORBEILLES;TOX
*****************************************************************}
unit PaieToxCtrl;

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
  StdCtrls,
  Spin,
  ExtCtrls,
  FileCtrl,
  ComCtrls,
  Registry,
  Buttons,
  Menus,
  ShellApi,
  DBTables,
  Hctrls,
  HmsgBox,
  uTob ,
  PaieTOXInfoToxZip,
  uToz,
  uToxClasses,
  ImgList,
{$IFNDEF EAGLSERVER}
  Fe_Main,
{$ENDIF}  
  ParamSoc,
  HStatus ;

type
  TFicheCtrlTox = class(TForm)
    Timer1: TTimer;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    PanelCR: TPanel;
    PanelCA: TPanel;
    PanelCE: TPanel;
    PanelCD: TPanel;
    PanelCT: TPanel;
    Bdelete: TPanel;
    Label8: TLabel;
    Scrute: TBitBtn;
    Panel11: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    CD: TEdit;
    CA: TEdit;
    CT: TEdit;
    CE: TEdit;
    CR: TEdit;
    Panel12: TPanel;
    Label1: TLabel;
    ScanTime: TSpinEdit;
    Label2: TLabel;
    ConfirmDel: TCheckBox;
    PopupMenuBin: TPopupMenu;
    Supprimerlefichierselectionne: TMenuItem;
    Supprimerlacorbeillecourante: TMenuItem;
    Supprimertouteslescorbeilles: TMenuItem;
    ImageList1: TImageList;
    BINA: TListView;
    BINT: TListView;
    BIND: TListView;
    BINE: TListView;
    BINR: TListView;
    N1: TMenuItem;
    Informations: TMenuItem;
    OthersIcons: TImageList;
    Filtres: THValComboBox;
    GroupBox1: TGroupBox;
    TriAlpha: TRadioButton;
    TriChrono: TRadioButton;
    Jauge: TCheckBox;
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ScruteClick(Sender: TObject);
    procedure FiltresChange(Sender: TObject);
    procedure FCDDblClick (T : TListView ; S : string);
    procedure FormCreate(Sender: TObject);
    procedure FCDClick(Sender: TObject);
    procedure DeleteCurrentFileClick (Sender: TObject);
    procedure DeleteCurrentBinClick (Sender: TObject);
    procedure DeleteAllClick (Sender: TObject);
    procedure OnPopup (Sender : TObject) ;
    procedure FCDMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FCDDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure FCDDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FCDMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure Purge(Sender : TObject) ;
    procedure PurgeCorbeille(AutoPurge, Fic : boolean ; Corbeille : TListView ; FileToDelete : string) ;
    procedure LaPurge (AutoPurge : boolean ; Enveloppes : TCollectionEnveloppes ; FileName : string) ;
    procedure BINDDblClick(Sender: TObject);
    procedure BINTDblClick(Sender: TObject);
    procedure BINADblClick(Sender: TObject);
    procedure BINEDblClick(Sender: TObject);
    procedure BINRDblClick(Sender: TObject);
    procedure InformationsClick(Sender: TObject);
    procedure ResizeBins(Sender: TObject) ;
    procedure LVColumnClick(Sender: TObject; Column: TListColumn);
    procedure LVCompare(Sender: TObject; Item1, Item2: TListItem; Data: Integer; var Compare: Integer);
  private
    procedure WMDROPFILES(var Message: TWMDROPFILES);  message WM_DROPFILES;
    procedure LireParam;
    procedure MakeItems (st : String ; TLV : TListView ; TLI : TListItem) ;
    procedure UpdateBins ;
    procedure Suppression(F: String);
    function  GetRep(T: TListView): string;
    { Déclarations privées }
  public
    { Déclarations publiques }
    ListItemsBIND, ListItemsBINA, ListItemsBINT, ListItemsBINE, ListItemsBINR : TListItem ;
    MasqueBIN : string ;
    FileCur, FileSup : TListView ;
    RepDir : String ;
    CanDragDrop, FInvertSort : boolean ;
    iDebBin, iFinBin : integer ;
  end;


// Filtre personnalisé des corbeilles d'échanges
Const FILTRECORBEILLETOXPERSO : String = 'DefaultFilter' ;
      FILTREPARDEFAUT         : String = '*.*' ;

//
// Procédures ou fonctions à exporter
//
procedure PaieAfficheCorbeille ;
procedure PaiePurgeLesCorbeillesTox ;
procedure PaieCountFilesDir( Path, Mask : string; var Files, Dirs , octets : int64; ScanSubDir : boolean = False );

implementation

uses PaieToxCtrlEnv, PaieTOXInfoTox, PaieToxConsultBis, HPanel, HEnt1, UIUtil ;

{$R *.DFM}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Affichage des corbeilles
Mots clefs ... : CORBEILLES;TOX;FO;
*****************************************************************}
procedure PaieAfficheCorbeille ;
var
  X  : TFicheCtrlTox;
  PP : THPanel ;
begin
  PP:=FindInsidePanel ;
  X := TFicheCtrlTox.Create ( Application ) ;
  if PP=Nil then
   BEGIN
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Scrutation des corbeilles toutes les x minutes
Mots clefs ... : CORBEILLE;TOX;FO;
*****************************************************************}
procedure TFicheCtrlTox.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False ;

  // Corbeilles Départ
  if (DirectoryExists (CD.Text)) and (CD.Text <> '') then
  begin
    if Not BIND.Visible then BIND.visible := True ;
    BIND.Hint := IntToStr ( BIND.Items.Count) + ' fichiers' ;
    BIND.ShowHint := True ;
  end else
  begin
     if BIND.Visible then BIND.visible := False ;
  end;

  // Corbeille Arrivée
  if (DirectoryExists (CA.Text)) and (CA.Text <> '') then
  begin
    if Not BINA.Visible then BINA.visible := True ;
    BINA.Hint := IntToStr ( BINA.Items.Count ) + ' fichiers' ;
    BINA.ShowHint := True ;
  end else
  begin
    if BINA.Visible then BINA.visible := False ;
  end;

  // Cobeille fichiers traités
  if (DirectoryExists (CT.Text)) and (CT.Text <> '') then
  begin
    if Not BINT.Visible then BINT.visible := True ;
    BINT.Hint := IntToStr ( BINT.items.Count ) + ' fichiers' ;
    BINT.ShowHint := True ;
  end else
  begin
    if BINT.Visible then BINT.visible := False ;
  end;

  // Corbeille fichiers émis
  if (DirectoryExists (CE.Text)) and (CE.Text <> '') then
  begin
    if Not BINE.Visible then BINE.visible := True ;
    BINE.Hint := IntToStr ( BINE.items.Count ) + ' fichiers' ;
    BINE.ShowHint := True ;
  end else
  begin
    if BINE.Visible then BINE.visible := False ;
  end;

  // Crbeille Rejet
  if (DirectoryExists (CR.Text)) and (CR.Text <> '') then
  begin
    if Not BINR.Visible then BINR.visible := True ;
    BINR.Hint := IntToStr ( BINR.items.Count ) + ' fichiers' ;
    BINR.ShowHint := True ;
  end else
  begin
    if BINR.Visible then BINR.visible := False ;
  end;

  Timer1.Enabled := True ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Déplacement d'un fichier TOX d'une corbeille à une autre
Mots clefs ... :
*****************************************************************}
procedure TFicheCtrlTox.WMDROPFILES(var Message: TWMDROPFILES);
var
  NumFiles : longint;
  i        : longint;
  buffer   : array[0..MAX_PATH] of char;
  F        : string ;
begin
 {How many files are being dropped}
  NumFiles := DragQueryFile(Message.Drop,$FFFFFFFF,nil,0);

 {Accept the dropped files}
  for i := 0 to (NumFiles - 1) do
    begin
    DragQueryFile(Message.Drop,i,@buffer,sizeof(buffer));

    if FileCur <> Nil then
      begin
      F := Buffer ;
      RepDir := GetRep ( FileCur ) ;
      CopyFile ( pchar(F), pchar(RepDir+ExtractFileName(F)), False) ;
      end ;
    end;

  DragFinish(Message.Drop);
  Timer1Timer(Nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Constitution des corbeilles
Mots clefs ... : BO;
*****************************************************************}
procedure TFicheCtrlTox.MakeItems (st : String ;TLV : TListView ; TLI : TListItem) ;
var TSR     : TsearchRec ;
    i       : integer ;
    masque, sousmasque, path : string ;
    Files, Dirs , octets : int64 ;
begin
  //
  // Suppression des items et icones des corbeilles
  //
  Application.ProcessMessages ;
  TLV.Items.BeginUpdate ; ;
  TLV.Items.Clear ;
  //
  // Extraction du chemin et du masque des fichiers à aficher
  //
  path   := ExtractFilePath(st) ;
  masque := ExtractFileName(st) ;
  Files := 0 ; Dirs := 0 ; octets := 0 ;
  if Jauge.Checked then
     BEGIN
     PaieCountFilesDir( Path, masque, Files, Dirs, octets, False ) ;
     InitMove (Files, 'Actualisation en cours ...') ;
     END else
     Files := 100 ;
  repeat
    BEGIN
    //
    // si le masque est composé, on récupère chaque sous-masque
    //
    sousmasque := ReadTokenSt(masque) ;
    if FindFirst (path + sousmasque, faArchive, TSR) = 0 then
       repeat
         TLI := TLV.Items.Add ;
         TLI.Caption := TSR.Name ;
         //
         // récupération des icônes associées aux fichiers
         //
         if UpperCase(ExtractFileExt(TLI.Caption)) = '.ENV' then
            i := 0 else
         if UpperCase(ExtractFileExt(TLI.Caption)) = '.TOX' then
            i := 1 else
         if UpperCase(ExtractFileExt(TLI.Caption)) = '.ZIP' then
            i := 5 else
            i := 6;
           { A gérer plus tard
            BEGIN
            TheIcon := TIcon.Create ;
            TheIcon.Handle := ExtractAssociatedIcon(Application.Handle, PChar(Path + TLI.Caption), TheWord) ;
            i := OthersIcons.AddIcon(TheIcon) ;
            TheIcon.Free ;
            END ;  }
         TLI.ImageIndex := i ;
         if Jauge.Checked then MoveCur(False) ;
       until FindNext (TSR) <> 0 ;
    FindClose (TSR) ;
    END ;
  until masque = '' ;
  if Jauge.Checked then FiniMove ;

  TLV.Items.EndUpdate ;
  if TLV.Items.Count > 0 then
     BEGIN
     TLV.Hint := IntToStr (TLV.Items.Count) + ' fichiers' ;
     TLV.ShowHint := True ;
     END ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Mise à jour du contenu des corbeilles
Mots clefs ... : BO;
*****************************************************************}
procedure TFicheCtrlTox.UpdateBins ;
begin
  MakeItems(IncludeTrailingBackslash(CD.Text) + MasqueBIN, BIND, ListItemsBIND) ;
  MakeItems(IncludeTrailingBackslash(CE.Text) + MasqueBIN, BINE, ListItemsBINE) ;
  MakeItems(IncludeTrailingBackslash(CA.Text) + MasqueBIN, BINA, ListItemsBINA) ;
  MakeItems(IncludeTrailingBackslash(CT.Text) + MasqueBIN, BINT, ListItemsBINT) ;
  MakeItems(IncludeTrailingBackslash(CR.Text) + MasqueBIN, BINR, ListItemsBINR) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Ajustement de la largeur des corbeilles en fonction de la
Suite ........ : taille de la fenêtre
Mots clefs ... : BO;
*****************************************************************}
procedure TFicheCtrlTox.ResizeBins(Sender: TObject) ;
var i : double ;
begin
  i := Int(TabSheet1.Width / 5) ;
  PanelCA.Width := StrToInt(FloatToStr(i)) ;
  PanelCT.Width := StrToInt(FloatToStr(i)) ;
  PanelCD.Width := StrToInt(FloatToStr(i)) ;
  PanelCE.Width := StrToInt(FloatToStr(i)) ;
  PanelCR.Width := StrToInt(FloatToStr(i)) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Chargement des paramètres et affichage des corbeilles
Mots clefs ... : CORBEILLES;TOX;FO;
*****************************************************************}
procedure TFicheCtrlTox.LireParam ;
var
  LesSites: TCollectionSites ;
begin
  LesSites:=TCollectionSites.Create(TCollectionSite,True) ;
  if LesSites.LeSiteLocal<>Nil then
    begin
    //
    // récupération des chemin d'accès des corbeilles
    //
    With LesSites.LeSiteLocal do
      begin
      CA.Text := SSI_CARRIVE ; // Q.FindField('SSI_CARRIVE').AsString ;
      CD.Text := SSI_CDEPART ; // Q.FindField('SSI_CDEPART').AsString ;
      CE.Text := SSI_CENVOYE ; // Q.FindField('SSI_CENVOYE').AsString ;
      CT.Text := SSI_CTRAITE ; // Q.FindField('SSI_CTRAITE').AsString ;
      CR.Text := SSI_CREJET ; // Q.FindField('SSI_CREJET').AsString ;
      //
      // mise à jour du contenu des corbeilles
      //
      UpdateBins ;

      Timer1.Enabled := False ;
      Timer1.Interval := ScanTime.Value * 60 * 1000 ;
      Timer1.Enabled := True ;
      end ;
    end
  else
    begin
    PGIBox ('Vous n''avez pas défini de site local dans la table des sites ! ', '');
    end;
  LesSites.Free ;
end;

procedure TFicheCtrlTox.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Timer1.Enabled := False ;
  Action := caFree ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Scrutation manuelle des corbeilles
Mots clefs ... : BO;
*****************************************************************}
procedure TFicheCtrlTox.ScruteClick(Sender: TObject);
begin
  //
  // mise à jour du contenu des corbeilles SANS épuration
  //
  UpdateBins ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Changement du filtres
Mots clefs ... :
*****************************************************************}
procedure TFicheCtrlTox.FiltresChange(Sender: TObject);
begin
  case Filtres.ItemIndex of
    0 : BEGIN
        MasqueBIN := '*_*.ENV' ;
        UpdateBins;
        END ;
    1 : BEGIN
        MasqueBIN := '*_*.TOX' ;
        UpdateBins;
        END ;
    2 : BEGIN
        MasqueBIN := '*_*.ZIP' ;
        UpdateBins;
        END ;
    3 : BEGIN
        MasqueBIN := FILTREPARDEFAUT ;
        UpdateBins ;
        END ;
    4 : BEGIN
        MasqueBIN := '*_*.*' ;
        UpdateBins;
        END ;
    5 : BEGIN
        // Chargement du filtre sauvegardé
{$IFNDEF EAGLSERVER}
        MasqueBIN := AGLLanceFiche('MBO','FILTREPERSOTOX','','',GetSynRegKey (FILTRECORBEILLETOXPERSO, FILTREPARDEFAUT, True)) ;
{$ENDIF}        
        if MasqueBIN <> '' then
           BEGIN
           SaveSynRegKey(FILTRECORBEILLETOXPERSO, MasqueBIN, True) ;  // sauvegarde du nouveau filtre
           UpdateBins ; ;  // actualisation des corbeilles
           END ;
        END ;
    end ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Consultation d'un fichier TOX
Mots clefs ... : CORBEILLES;TOX;FO;
*****************************************************************}
procedure TFicheCtrlTox.FCDDblClick (T : TListView ; S : string);
var MaToz : Toz ;
    st    : string ;
    Okay  : boolean ;
begin
  if T.selected <> nil then
    begin
    CanDragDrop := False ;
    Okay := False ;
    MaToz := Nil ;
    if      UpperCase ( ExtractFileExt ( T.Selected.Caption ) ) = '.ENV' then
      DisplayEnveloppe ( S + T.Selected.Caption)
    else if UpperCase ( ExtractFileExt ( T.Selected.Caption ) ) = '.TOX' then
      PaieConsulteAutoTox  ( S + T.Selected.Caption, True)
    else if UpperCase ( ExtractFileExt ( T.Selected.Caption ) ) = '.ZIP' then
      begin
      //
      // Extraction du fichier .TOX du .ZIP
      TRY
        TRY
        MaToz := TOZ.Create ;
        if MaToz.OpenZipFile ( S + T.Selected.Caption, moOpen ) then
          begin
          if MaToz.OpenSession (osExt) then
            begin
            if MaToz.SetDirOut (S) then
              begin
              if MaToz.CloseSession then
                begin
                // Suppression de la TOZ
                Okay := True ;
                MaToz.Free ;
                MaToz := Nil ;
                end ;
              end ;
            end ;
          end ;
        EXCEPT
          on Erreur : ETozErreur do
            begin
            ShowMessage ( Erreur.Message ) ;
            Okay := False ;
            end ;
        END ;
      FINALLY
        if MaToz <> Nil then MaToz.Free ;
      END ;
      if Okay then // l'extraction s'est bien passée
        begin
        //
        // Consultation du fichier .TOX
        //
        st := Copy ( S + T.Selected.Caption, 1, Length ( S + T.Selected.Caption ) - 4 ) ;
        PaieConsulteAutoTox  ( st, True) ;
        //
        // Suppression du fichier .TOX
        //
        DeleteFile ( st ) ;
        end ;
      end ;
    end ;
end;

procedure TFicheCtrlTox.FormCreate(Sender: TObject);
begin
  // Filtre des corbeilles par défauts'il existe
  MasqueBIN := GetSynRegKey (FILTRECORBEILLETOXPERSO, FILTREPARDEFAUT, True) ;
  RepDir  := '' ;
  FileCur := Nil ;
  LireParam ;
  Self.FormStyle := fsNormal ;
  // Indique à Windows que cette forme accepte le "déposes" de fichier
  DragAcceptFiles(Handle, True);
  CanDragDrop := True ;
  FInvertSort := False ;
end;

procedure TFicheCtrlTox.FCDClick (Sender : TObject);
begin
  RepDir  := GetRep ( TListView ( Sender ) ) ;
  FileCur := TListView ( Sender ) ;
  FileSup := FileCur ;
  CanDragDrop := True ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Suppression d'un fichier TOX
Mots clefs ... : SUPPRESSION TOX
*****************************************************************}
procedure TFicheCtrlTox.Suppression ( F : String ) ;
var
  R: Integer ;
begin
  if not ConfirmDel.Checked then R := mrYes
  else R := MessageDlg ( 'Confirmez-vous la suppression du fichier ' + RepDir + F + ' ?', mtConfirmation, [mbYes, mbNo], -1 ) ;
  if R = mrYes then DeleteFile ( RepDir + F ) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Suppression du fichier TOX sélectionné
Mots clefs ... :
*****************************************************************}
procedure TFicheCtrlTox.DeleteCurrentFileClick (Sender : TObject);
begin
  RepDir := '' ;
  if (FileCur <> Nil) then
    PurgeCorbeille (True, True, FileCur, ExtractFileName (FileCur.Selected.Caption)) ;   Timer1Timer ( Nil ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Suppression des fichiers de la corbeille courante
Mots clefs ... : 
*****************************************************************}
procedure TFicheCtrlTox.DeleteCurrentBinClick (Sender : TObject);
var
  iCompteur : integer ;
begin
  if (FileCur <> Nil) and (FileCur.Items.Count > 0) then
    begin
    FileCur.Items.BeginUpdate ;
    For iCompteur:=0 to FileCur.Items.Count-1 do
      begin
      Suppression ( FileCur.Items.Item[0].Caption ) ;
      Filecur.Items.Delete (0) ;
      end ;
    FileCur.Items.EndUpdate ;
    Timer1Timer ( Nil ) ;
    end ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Suppression de tous les fichiers TOX
Mots clefs ... : 
*****************************************************************}
procedure TFicheCtrlTox.DeleteAllClick (Sender : TObject);
begin
  RepDir := IncludeTrailingBackslash(CD.Text) ;
  FileCur := BIND ;
  DeleteCurrentBinClick ( Nil ) ;
  RepDir := IncludeTrailingBackslash(CA.Text) ;
  FileCur := BINA ;
  DeleteCurrentBinClick ( Nil ) ;
  RepDir := IncludeTrailingBackslash(CT.Text) ;
  FileCur := BINT ;
  DeleteCurrentBinClick ( Nil ) ;
  RepDir := IncludeTrailingBackslash(CE.Text) ;
  FileCur := BINE ;
  DeleteCurrentBinClick ( Nil ) ;
  RepDir := IncludeTrailingBackslash(CR.Text) ;
  FileCur := BINR ;
  DeleteCurrentBinClick ( Nil ) ;
  Timer1Timer ( Nil ) ;
  RepDir := '' ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Contrôle de l'affichage du Popup des corbeilles
Mots clefs ... : FO;
*****************************************************************}
procedure TFicheCtrlTox.OnPopup (Sender : Tobject) ;
begin
  Popupmenubin.Items[4].Enabled := True ;
  if FileCur.Items.Count = 0 then
    begin
    Popupmenubin.Items[2].Enabled := False ;
    Popupmenubin.Items[3].Enabled := False ;
    Popupmenubin.Items[0].Enabled := False ;
    end else
    begin
    if (FileCur.Selected = nil) then
      begin
      Popupmenubin.Items[2].Enabled := False ;
      Popupmenubin.Items[3].Enabled := True ;
      Popupmenubin.Items[0].Enabled := False ;
      end else
      begin
      Popupmenubin.Items[2].Enabled := True ;
      Popupmenubin.Items[3].Enabled := True ;
      if pos (UpperCase (ExtractFileExt (FileCur.Selected.Caption )),'.ZIP.TOX') >0 then
        Popupmenubin.Items[0].Enabled := True else
        Popupmenubin.Items[0].Enabled := False ;
      end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Déplacement d'un fichier
Mots clefs ... : CORBEILLES;TOX;
*****************************************************************}
procedure TFicheCtrlTox.FCDMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if TListView(Sender).Selected <> nil then
    if Button = mbLeft then
      if CanDragDrop then TListView(Sender).BeginDrag (False, 10) ;
end;

procedure TFicheCtrlTox.FCDDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  if ( Sender is TListView ) and ( Source is TListView ) and ( Sender <> Source ) then Accept := True
  else Accept := False ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Déplacement de fichiers d'une corbeille à une autre
Mots clefs ... :
*****************************************************************}
procedure TFicheCtrlTox.FCDDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  Tsrc, Tdest          : TListView ;
  RepSrc, AttachedFile : string ;

begin
  Tsrc  := TListView(Source) ;  // corbeille source
  Tdest := TListView(Sender) ;  // corbeille destination
  if Tsrc.Selected.Caption <> '' then
      begin
      RepSrc := GetRep (TSrc)    ;  // on stocke le répertoire source
      RepDir := GetRep ( Tdest ) ;  // on stocke le répertoire destination

      // si le fichier à déplacer est un .ENV
      if UpperCase (ExtractFileExt (TSrc.Selected.Caption)) = '.ENV' then
         begin
         // existe-t-il un .TOX associé ?
         if FileExists (RepSrc + ChangeFileExt (TSrc.Selected.Caption, '.TOX')) then
            AttachedFile := ChangeFileExt (TSrc.Selected.Caption, '.TOX') else
         // ou existe-t-il un .ZIP associé ?
         if FileExists (RepSrc + ChangeFileExt (TSrc.Selected.Caption, '.TOX') + '.ZIP') then
            AttachedFile := ChangeFileExt (TSrc.Selected.Caption, '.TOX') + '.ZIP' ;
         end ;

      // si le fichier à déplacer est un .TOX
      if UpperCase (ExtractFileExt (TSrc.Selected.Caption)) = '.TOX' then
         // existe-t-il un .ENV associé ?
         if FileExists (RepSrc + ChangeFileExt (TSrc.Selected.Caption, '.ENV')) then
            AttachedFile := ChangeFileExt (TSrc.Selected.Caption, '.ENV') ;

      // si le fichier à déplacer est un .ZIP
      if UpperCase (ExtractFileExt (TSrc.Selected.Caption)) = '.ZIP' then
         // existe-t-il un .ENV associé ?
         if FileExists (RepSrc + ChangeFileExt (ChangeFileExt (TSrc.Selected.Caption, ''), '.ENV')) then
            AttachedFile := ChangeFileExt (ChangeFileExt (TSrc.Selected.Caption, ''), '.ENV') ;

      // on copie le fichier sélectionné dans la corbeille de destination
      if CopyFile ( pchar(RepSrc + TSrc.Selected.Caption), pchar(RepDir + TSrc.Selected.Caption), False) then
         // et on le supprime de la corbeille source
         if DeleteFile(pchar(RepSrc + TSrc.Selected.Caption)) then
            BEGIN
            Tdest.Items.add ;
            Tdest.Items.Item[Tdest.Items.Count-1].Caption := TSrc.Selected.Caption ;
            Tsrc.Items.Delete (FileCur.FindCaption (0, TSrc.Selected.Caption, False, True, False).Index) ;
            END ;
      // on copie le fichier attaché dans la corbeille de destination
      if CopyFile ( pchar(RepSrc + AttachedFile), pchar(RepDir + AttachedFile), False) then
         // et on le supprime de la corbeille source
         if DeleteFile(pchar(RepSrc + AttachedFile)) then
            //
            // si l'extension du fichier attaché fait partie du masque d'affichage de la corbeille
            //
            if Pos(ExtractFileExt(AttachedFile),MasqueBIN)<>0 then
               BEGIN
               Tdest.Items.add ;
               Tdest.Items.Item[Tdest.Items.Count-1].Caption := AttachedFile ;
               Tsrc.Items.Delete (FileCur.FindCaption (0, AttachedFile, False, True, False).Index) ;
               END ;

      MakeItems(RepSrc + MasqueBIN, Tsrc , TListItem(Tsrc.Items)) ;
      MakeItems(RepDir + MasqueBIN, Tdest, TListItem(Tdest.Items)) ;
      Timer1Timer(Nil);
      end ;
end;

function TFicheCtrlTox.GetRep ( T : TListView ) : string ;
begin
    if T.Name = 'BINA' then Result := IncludeTrailingBackslash(CA.Text) else
    if T.Name = 'BINT' then Result := IncludeTrailingBackslash(CT.Text) else
    if T.Name = 'BIND' then Result := IncludeTrailingBackslash(CD.Text) else
    if T.Name = 'BINE' then Result := IncludeTrailingBackslash(CE.Text) else
    if T.Name = 'BINR' then Result := IncludeTrailingBackslash(CR.Text) else
       result := '' ;
end ;

procedure TFicheCtrlTox.FCDMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  if TListView(Sender).Color = clInfoBk then Exit ;
  if BIND.Enabled then BIND.Color := clWhite ;
  if BINA.Enabled then BINA.Color := clWhite ;
  if BINT.Enabled then BINT.Color := clWhite ;
  if BINE.Enabled then BINE.Color := clWhite ;
  if BINR.Enabled then BINR.Color := clWhite ;
  TListView(Sender).Color := clInfoBk ;
  FileCur := TListView(Sender) ;
  RepDir := GetRep ( FileCur ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Balayage des corbeilles pour la purge
Mots clefs ... : PURGE;TOX;
*****************************************************************}
procedure TFicheCtrlTox.Purge(Sender : TObject) ;
begin
  ResizeBins(Sender) ;
  FileCur := BINA ;
  PurgeCorbeille (False, False, BINA, '') ;
  FileCur := BINR ;
  PurgeCorbeille (False, False, BINR, '') ;
  FileCur := BIND ;
  PurgeCorbeille (False, False, BIND, '') ;
  FileCur := BINT ;
  PurgeCorbeille (True, False, BINT, '') ;    
  FileCur := BINE ;
  PurgeCorbeille (True, False, BINE, '') ;
  RepDir := '' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Traitement des différents cas pour la purge
Mots clefs ... : PURGE;TOX;ZIP;
*****************************************************************}
procedure TFicheCtrlTox.PurgeCorbeille(AutoPurge, Fic : boolean ; Corbeille : TListView ; FileToDelete : string) ;
var
  Enveloppes: TCollectionEnveloppes ;
  FileName : string ;
  BType : boolean ;
begin
  Enveloppes:=TCollectionEnveloppes.Create(TCollectionEnveloppe) ;
  //
  // définition des bornes de la boucle de suppression
  //
  iDebBin := 1 ;
  if FileToDelete = '' then  // si pas de fichier spécifié on purge la corbeille entière
     iFinBin := Corbeille.Items.Count else
     iFinBin := 1 ; // on ne supprime que le fichier

  //
  // traitement de la suppresion
  //
  While (iDebBin <= iFinBin) and (Corbeille.Items.Count > 0) do
    begin
    BType := True ;
    if Corbeille.Items.Count = 0 then break ;
      RepDir := GetRep (Corbeille) ;
      if FileToDelete = '' then FileName := Corbeille.Items.Item[iDebBin-1].Caption else FileName := FileToDelete ;
      //
      // si le fichier ne contient pas d'underscore ce n'est pas un fichier "toxé" donc suppression
      //
      if ((Pos('_', FileName)<> 0) AND
          (UpperCase (ExtractFileExt (FileName)) <> '.ENV') AND
          (UpperCase (ExtractFileExt (FileName)) <> '.TOX') AND
          (UpperCase (ExtractFileExt (FileName)) <> '.ZIP')) OR (Pos('_', FileName)=0) then
         begin
         if DeleteFile (RepDir + FileName) then
            begin
            FileCur.Items.Delete (FileCur.FindCaption (0, FileName, False, True, False).Index) ;
            Dec(iFinBin) ;
            end ;
         end else
         //
         // Purge Automatique
         //
         BEGIN
         if Fic then BType := False ; // si on efface seulement un fichier alors pas de purge automatique
         if AutoPurge then
            begin
            // CAS DU FICHIER .ENV
            if UpperCase (ExtractFileExt (FileName)) = '.ENV' then
             //  if (Corbeille <> BINA) and (Corbeille <> BIND) then
                   LaPurge (BType, Enveloppes, RepDir + FileName) else

            // CAS DU FICHIER .TOX
            if (UpperCase (ExtractFileExt (FileName)) = '.TOX') then
                begin
                // on recherche si le .ENV auquel le .TOX est attaché existe
                if FileExists (RepDir + ChangeFileExt (FileName, '.ENV')) then
                   FileName := ChangeFileExt (FileName, '.ENV') ; //else
                // si le .ENV n'existe pas , la purge n'est pas lancée sur les corbeille d'arrivée et de départ
                   if (Corbeille <> BINA) and (Corbeille <> BIND) then
                      LaPurge (BType, Enveloppes, RepDir + FileName) ;
                end else
            // CAS DU FICHIER .ZIP
            if (UpperCase (ExtractFileExt (FileName)) = '.ZIP') then
                begin
                if UpperCase (ExtractFileExt (FileName)) = '.ZIP' then
                   FileName := ChangeFileExt (FileName, '') ; // on obtient le fichier .TOX

                // on recherche si le .ENV auquel le .TOX est attaché existe
                if FileExists (RepDir + ChangeFileExt (FileName, '.ENV')) then
                   FileName := ChangeFileExt (FileName, '.ENV') else
                   FileName := FileName + '.ZIP' ;
                // si le .ENV existe, la purge est lancée
                if ( UpperCase (ExtractFileExt (FileName)) = '.ENV') or
                   ((UpperCase (ExtractFileExt (FileName)) = '.ZIP') and (Corbeille <> BINA) and (Corbeille <> BIND)) then
                   LaPurge (BType, Enveloppes, RepDir + FileName) ;
                end else
                Inc(iDebBin) ;
            end else
            Inc(iDebBin) ;
         END ;
      end ;
  Enveloppes.Free ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Purge une corbeille :
Suite ........ :   - des fichiers .ENV, .TOX et .TOX.ZIP de 7 jours ou plus
Suite ........ :   - des autres fichiers "parasites"
Mots clefs ... : CORBEILLES;PURGE;TOX;ZIP;FO;BO;
*****************************************************************}
procedure TFicheCtrlTox.LaPurge (AutoPurge : boolean ; Enveloppes : TCollectionEnveloppes ; FileName : string) ;
var i, iAnciennete : integer ;
begin
  Enveloppes.Clear ;
  Enveloppes.LoadEnveloppe(FileName) ;
  if AutoPurge then iAnciennete := GetParamSoc('SO_GCTOXEPURE') else iAnciennete := -1 ;
  if Int( Now - Enveloppes.Items[0].DateMsg ) >= iAnciennete then
     begin
     for i := 0 to Enveloppes.Items[0].LesFichiers.Count - 1 do
         begin
         if DeleteFile (RepDir + Enveloppes.Items[0].LesFichiers.Items[i].FileName) then
            begin
            FileCur.Items.Delete (FileCur.FindCaption (0, Enveloppes.Items[0].LesFichiers.Items[i].FileName, False, True, False).Index) ;
            Dec(iFinBin) ;
            end ;
         end ;
     if DeleteFile( FileName ) then
        begin
        FileCur.Items.Delete(FileCur.FindCaption (0,extractfilename(filename) , False, True, False).Index) ;
        Dec(iFinBin) ;
        end ;
     end else Inc(iDebBin) ;
  Enveloppes.Clear ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Balayage des corbeilles de la TOX pour la suppression des
Suite ........ : fichiers présents depuis plus de 7 jours sans affichage.
Mots clefs ... : PURGE;TOX;
*****************************************************************}
procedure PaiePurgeLesCorbeillesTox ;
Var XX : TFicheCtrlTox;
BEGIN
XX := TFicheCtrlTox.Create ( Application ) ;
XX.Visible := False ;
try
  XX.Purge(Nil) ;
 finally
  XX.Free ;
 end ;
END ;


procedure TFicheCtrlTox.BINDDblClick(Sender: TObject);
begin
  FCDDblClick (BIND, IncludeTrailingBackslash(CD.Text)) ;
end;

procedure TFicheCtrlTox.BINTDblClick(Sender: TObject);
begin
  FCDDblClick (BINT, IncludeTrailingBackslash(CT.Text)) ;
end;

procedure TFicheCtrlTox.BINADblClick(Sender: TObject);
begin
  FCDDblClick (BINA, IncludeTrailingBackslash(CA.Text)) ;
end;

procedure TFicheCtrlTox.BINEDblClick(Sender: TObject);
begin
  FCDDblClick (BINE, IncludeTrailingBackslash(CE.Text)) ;
end;

procedure TFicheCtrlTox.BINRDblClick(Sender: TObject);
begin
  FCDDblClick (BINR, IncludeTrailingBackslash(CR.Text)) ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Affiche les informations de fichier :
Suite ........ : date, taille, ...
Mots clefs ... : FO;BO;
*****************************************************************}
procedure TFicheCtrlTox.InformationsClick(Sender: TObject);
var MaToz : Toz ;
    MaTob : Tob ;
    Okay  : boolean ;
    F     : file of byte ;
    iSize : integer ;
begin
  if FileCur.Selected <> nil then
    begin
    MaToz := Nil ;
    MaTob := Nil ;
    Okay  := False ;

    //le fichier est un .ZIP
    if UpperCase ( ExtractFileExt (FileCur.Selected.Caption) ) = '.ZIP' then
      begin
      TRY
        TRY
        MaToz := TOZ.Create ;
        if MaToz.OpenZipFile ( RepDir + FileCur.Selected.Caption, moOpen ) then
          begin
          MaTob := Tob.Create('TobToz',Nil,-1) ;
          MaTob := MaToz.ConvertListInTob ;
          Okay  := True ;
          MaToz.Free ;
          MaToz := Nil ;
          end ;
        EXCEPT
          on Erreur : ETozErreur do
            begin
            ShowMessage ( Erreur.Message ) ;
            Okay := False ;
            end ;
        END ;
      FINALLY
        if MaToz <> Nil then MaToz.Free ;
      END ;
      if Okay then
        begin
        InfoToxZip ( MaTob );
        MaTob.Free ;
        end;
      end else
      // le fichier est un .TOX
      if UpperCase (ExtractFileExt (FileCur.Selected.Caption) ) = '.TOX' then
        begin
        AssignFile (F, RepDir + FileCur.Selected.Caption);
        Reset(F);
        iSize := FileSize(F);
        CloseFile (F) ;
        InfoTox ( ExtractFileName (FileCur.Selected.Caption), iSize ) ;
        end ;
  end ;
end;

procedure TFicheCtrlTox.LVCompare(Sender: TObject; Item1, Item2: TListItem; Data: Integer; var Compare: Integer);
begin
  Compare := CompareText(Item1.Caption,Item2.Caption) ;
  if FInvertSort then Compare := -Compare;
end;

procedure TFicheCtrlTox.LVColumnClick(Sender: TObject; Column: TListColumn);
begin
  FInvertSort := not FInvertSort ;
  TListView(Sender).CustomSort(nil,0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dev
Créé le ...... : 09/04/2003
Modifié le ... : 09/04/2003
Description .. : Compte le nombre de fichiers, de dossiers et de bytes 
Suite ........ : occupés par les fichiers contenus dans un répertoire donné
Suite ........ : et pour un type de fichiers donné. 
Suite ........ : Le comptage peut éventuellement inclure les sous-dossiers.
Mots clefs ... : FICHIERS
*****************************************************************}
procedure PaieCountFilesDir( Path, Mask : string; var Files, Dirs , octets : int64; ScanSubDir : boolean = False );
var RecFileInfo: TSearchRec;
begin
  // Formate le path en supprimant le '\' final
  Path := ExcludeTrailingBackslash(Path);
  try
    // Quelque chose à compter ?
    If FindFirst( Path + '\' + Mask , faAnyFile , RecFileInfo ) = 0 then
       BEGIN
       // Laisser respirer
       //Application.ProcessMessages;
       // Démarre
       repeat
       // Elimine . et ..
       If ( RecFileInfo.Name <> '.' ) and ( RecFileInfo.Name <> '..' ) then
          BEGIN
          // On est sur un dossier
          If ( RecFileInfo.Attr and faDirectory ) = faDirectory then
             BEGIN
             Dirs := Dirs + 1 ;
             // Comptage des sous-dossiers demandé: on repart
             If ScanSubDir then
                PaieCountFilesDir( Path + '\' + RecFileInfo.Name , Mask , Files , Dirs , octets, True );
             END else
             BEGIN
             // On est sur un fichier : total fichiers + total bytes
             Files := Files + 1;
             //octets := octets + RecFileInfo.Size;
	     END ;
          END ;
       until FindNext( RecFileInfo ) <> 0;
       END ;
  finally
  FindClose( RecFileInfo );
  end;
end;

end.
