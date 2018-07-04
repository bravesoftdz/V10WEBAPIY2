unit uCPGedFileViewer;

interface

uses
{$IFDEF EAGLCLIENT}
  MainEagl,      // AGLLanceFiche
  UtileAGL,
{$ELSE}
  Db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_main,       // AGLLanceFiche
{$ENDIF}
  Windows,
  Messages,
  SysUtils,
  Variants,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  HSysMenu,
  Hctrls,
  Menus,
  ImgList,
  HImgList,
  ExtCtrls,
  HPanel,
  HTB97,
  uTob,
  uGedFileViewer; // TGedInternalViewer

type
  TFCPGetFileViewer = class(TForm)
    Dock: TDock97;
    ToolWindow: TToolWindow97;
    btnHelp: TToolbarButton97;
    btnAnnul: TToolbarButton97;
    btnAccept: TToolbarButton97;
    btnOpen: TToolbarButton97;
    bSauve: TToolbarButton97;
    btnCopy: TToolbarButton97;
    btnPrint: TToolbarButton97;
    btnPan: TToolbarButton97;
    btnZoom: TToolbarButton97;
    btnFitToPage: TToolbarButton97;
    btnFitToLarge: TToolbarButton97;
    btnFitTo100: TToolbarButton97;
    btnFirstPage: TToolbarButton97;
    btnPagePrevious: TToolbarButton97;
    btnPageNext: TToolbarButton97;
    btnLastPage: TToolbarButton97;
    panelInfoPages: THPanel;
    PanelView: THPanel;
    ImageListZoom: THImageList;
    PopupMenuZoom: THPopupMenu;
    ZoomForward: THMenuItem;
    ZoomBackward: THMenuItem;
    OpenDialog: THOpenDialog;
    HMTrad: THSystemMenu;
    BRotation90: TToolbarButton97;
    PopUpRevInteg: THPopupMenu;
    DOCUMENTATIONTRAVAUX: TMenuItem;
    MEMOCYCLE: TMenuItem;
    MEMOOBJECTIF: TMenuItem;
    MEMOSYNTHESE: TMenuItem;
    MEMOMILLESIME: TMenuItem;
    MEMOCOMPTE: TMenuItem;
    ToolbarButton971: TToolbarButton97;
    INFOCOMP: TMenuItem;

    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BRotation90Click(Sender: TObject);
    procedure bSauveClick(Sender: TObject);
    procedure btnCopyClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnZoomClick(Sender: TObject);
    procedure btnFitToPageClick(Sender: TObject);
    procedure btnFitToLargeClick(Sender: TObject);
    procedure btnFitTo100Click(Sender: TObject);
    procedure btnPanClick(Sender: TObject);
    procedure btnFirstPageClick(Sender: TObject);
    procedure btnPagePreviousClick(Sender: TObject);
    procedure btnPageNextClick(Sender: TObject);
    procedure btnLastPageClick(Sender: TObject);
    procedure ZoomForwardClick(Sender: TObject);
    procedure ZoomBackwardClick(Sender: TObject);
    procedure OnUpdateButtons(Sender: TObject; ButtonStates: TGedButtonStates);
    procedure INFOCOMPClick(Sender: TObject);
    procedure MEMOCYCLEClick(Sender: TObject);
    procedure MEMOOBJECTIFClick(Sender: TObject);
    procedure MEMOSYNTHESEClick(Sender: TObject);
    procedure MEMOMILLESIMEClick(Sender: TObject);
    procedure MEMOCOMPTEClick(Sender: TObject);
    procedure DOCUMENTATIONTRAVAUXClick(Sender: TObject);
    procedure PopUpRevIntegPopup(Sender: TObject);

  private
    { Déclarations privées }
    FEtatBoutons      : TGedButtonStates ;

    procedure ActivationBouton;

  public
    { Déclarations publiques }

    FTobEcr            : Tob;
    FTobGene           : Tob;
    FStFileName        : string;
    FGedInternalViewer : TGedInternalViewer ;
  end;

var FCPGetFileViewer: TFCPGetFileViewer;

procedure CPLanceFiche_CPGedFileViewer( vStFileName: string; vTobEcr : Tob);

implementation

uses uIUtil,       // PrepareInside
     uLibWindows,  // Rotate90
     uLibExercice, // CtxExercice
     uLibEcriture, // CSelectDBTobCompl
     SaisUtil,     // TOBM
     SaisComp,     // R_COMP
     uSmoothImage, // TSmoothImage
     {$IFDEF MODENT1}
     CPTypeCons,   // RcReviseur
     {$ELSE}
     {$ENDIF}
     CPRevDocTravaux_Tof, // CPLanceFiche_CPRevDocTravaux
     Hent1,        // CtxPcl
     Ent1;         // VH^.

{$R *.dfm}

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_CPGedFileViewer( vStFileName: string; vTobEcr : Tob);
var lCPGetFileViewer : TFCPGetFileViewer;
    PHRien : THPanel;
begin
  lCPGetFileViewer := TFCPGetFileViewer.Create(Application);

  lCPGetFileViewer.FTobEcr     := vTobEcr;
  lCPGetFileViewer.FStFileName := VStFileName;

  //if PrepareInside then
  //begin
  //  PHRien := FindInsidePanel;

  //  if PHRien <> nil then
  //  begin
  //    InitInside(lCPGetFileViewer, PHRien);
  ///    lCPGetFileViewer.Visible := True;
  //    lCPGetFileViewer.Show;
      //lCPGetFileViewer.ShowFile(FileName);
  //  end;
  //end
  //else
  //begin
    //if Modal then
    //begin
      lCPGetFileViewer.ShowModal;
      lCPGetFileViewer.Free;
    //end
    //else
  //end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCPGetFileViewer.FormCreate(Sender: TObject);
begin
  FGedInternalViewer := TGedInternalViewer.Create( PanelView ) ;
  FGedInternalViewer.OnUpdateButtons := OnUpdateButtons;

  FTobGene := Tob.Create('GENERAUX', nil, -1);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFCPGetFileViewer.FormShow(Sender: TObject);
begin
  if FStFileName <> '' then
  begin
    FGedInternalViewer.ShowFile(FStFileName);
    ActivationBouton;
  end;

  if (FTobEcr <> nil) and (FTobEcr.Getstring('E_GENERAL') <> '') then
  begin
    FTobGene.SelectDB( FTobEcr.GetString('E_GENERAL'), nil, False);
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFCPGetFileViewer.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FGedInternalViewer.Free;
  FTobGene.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFCPGetFileViewer.BRotation90Click(Sender: TObject);
var lBitMap : TBitMap;
begin
  lBitMap := TSmoothImage( PanelView.Controls[0] ).Picture;
  if lBitMap <> nil then
  begin
    lBitMap := Rotate90( lBitMap );
    TSmoothImage( PanelView.Controls[0] ).Picture := lBitmap;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.bSauveClick(Sender: TObject);
begin
  FGedInternalViewer.SaveAs;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnCopyClick(Sender: TObject);
begin
  FGedInternalViewer.Copy;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnPrintClick(Sender: TObject);
begin
  FGedInternalViewer.Print;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnPanClick(Sender: TObject);
begin
  FGedInternalViewer.Pan;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnZoomClick(Sender: TObject);
begin
  if FGedInternalViewer.ZoomwardMode = gzmZoomForward then
    FGedInternalViewer.ZoomForward
  else
    FGedInternalViewer.ZoomBackward;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnFitToPageClick(Sender: TObject);
begin
  FGedInternalViewer.FitToPage;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnFitToLargeClick(Sender: TObject);
begin
  FGedInternalViewer.FitToLarge;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnFitTo100Click(Sender: TObject);
begin
  FGedInternalViewer.FitTo100;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnFirstPageClick(Sender: TObject);
begin
  FGedInternalViewer.FirstPage;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnPagePreviousClick(Sender: TObject);
begin
  FGedInternalViewer.PagePrevious;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnPageNextClick(Sender: TObject);
begin
  FGedInternalViewer.PageNext;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.btnLastPageClick(Sender: TObject);
begin
  FGedInternalViewer.LastPage;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.ZoomForwardClick(Sender: TObject);
begin
  FGedInternalViewer.ZoomBackward;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.ZoomBackwardClick(Sender: TObject);
begin
  FGedInternalViewer.ZoomForward;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.OnUpdateButtons(Sender: TObject; ButtonStates: TGedButtonStates);
begin
  FEtatBoutons := ButtonStates ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.ActivationBouton;
begin
  BRotation90.Enabled := (PanelView.Controls[0] is TSmoothImage);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.INFOCOMPClick(Sender: TObject);
var
  OBM: TOBM;
  RC: R_COMP;
  AA: TActionFiche;
  ModBN : Boolean;
  lQEcr: TQuery;
begin
  lQEcr := nil;
  OBM   := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ECRITURE WHERE ' +
                    '(E_JOURNAL = "' + FTobEcr.GetValue('E_JOURNAL') + '") AND ' +
                    '(E_EXERCICE = "' + FTobEcr.GetValue('E_EXERCICE') + '") AND ' +
                    '(E_DATECOMPTABLE = "' + USDateTime(FTobEcr.GetValue('E_DATECOMPTABLE')) + '") AND ' +
                    '(E_NUMEROPIECE = ' + IntToStr(FTobEcr.GetValue('E_NUMEROPIECE')) + ') AND ' +
                    '(E_NUMLIGNE = ' + IntToStr(FTobEcr.GetValue('E_NUMLIGNE')) + ') AND ' +
                    '(E_NUMECHE = ' + IntToStr(FTobEcr.GetValue('E_NUMECHE')) + ') AND ' +
                    '(E_QUALIFPIECE = "' + FTobEcr.GetValue('E_QUALIFPIECE') + '")', False);

    // GCO - 26/07/2007 - FQ 21084
    if (CtxExercice.QuelExoDate( FTobEcr.GetString('E_EXERCICE')).EtatCpta <> 'OUV') then
      AA := TaConsult
    else
      AA := taModif;

    OBM := TOBM.Create(EcrGen, '', False);
    OBM.ChargeMvt( lQEcr );
    Ferme( lQEcr );

    RC.TOBCompl := nil ;

    if OBM <> nil then
    begin
      if FTobGene.GetValue('G_CUTOFF') = 'X' then
       begin
        RC.StLibre     := '---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
        RC.StComporte  := 'XXXXXXXXXX' ;
        RC.CutOffPer   := FTobGene.GetValue('G_CUTOFFPERIODE') ;
        RC.CutOffEchue := FTobGene.GetValue('G_CUTOFFECHUE') ;
        RC.TOBCompl    := CSelectDBTOBCompl(OBM,nil) ;
       end
        else
         begin
          RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
          RC.StComporte := 'XXXXXXXXXX';
         end ;
      ModBN := True;
      RC.Conso := True;
      RC.Attributs := False;
      RC.MemoComp := nil;
      RC.Origine := -1;
      if SaisieComplement(TOBM(FTobEcr), EcrGen, AA, ModBN, RC, False, True) then
      begin
        CMAJTOBCompl(OBM) ; // FB 19843
        FTobEcr.SetAllModifie(true) ;
        FTobEcr.UpdateDb;
        if FTobGene.GetValue('G_CUTOFF') = 'X' then
         RC.TOBCompl.InsertOrUpdateDB(false);
      end;
    end;

  finally
    Ferme(lQEcr);
    FreeAndNil(RC.TOBCompl) ;
    FreeAndNil(OBM);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.MEMOCYCLEClick(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( FTobEcr.GetString('E_GENERAL'), '', VH^.EnCours.Code, 0 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.MEMOOBJECTIFClick(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( FTobEcr.GetString('E_GENERAL'), '', VH^.EnCours.Code, 1 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.MEMOSYNTHESEClick(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( FTobEcr.GetString('E_GENERAL'), '', VH^.EnCours.Code, 2 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.MEMOMILLESIMEClick(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( FTobEcr.GetString('E_GENERAL'), '', VH^.EnCours.Code, 3 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.MEMOCOMPTEClick(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( FTobEcr.GetString('E_GENERAL'), '', VH^.EnCours.Code, 4 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.PopUpRevIntegPopup(Sender: TObject);
begin
  DOCUMENTATIONTRAVAUX.Visible := VH^.OkModRic and (VH^.Revision.Plan <> '');
end;

////////////////////////////////////////////////////////////////////////////////
procedure TFCPGetFileViewer.DOCUMENTATIONTRAVAUXClick(Sender: TObject);
var lStCycleRevision : string;
begin
  lStCycleRevision := FTobGene.GetString('G_CYCLEREVISION');

  MemoCycle.Visible := JaiLeRoleCompta( RcReviseur ) and
                       (VH^.Revision.Plan <> '') and
                       (lStCycleRevision <> '');

  MemoObjectif.Visible := (CtxPcl In V_Pgi.PgiContexte) and
                          JaiLeRoleCompta( RcReviseur ) and
                          (VH^.Revision.Plan <> '') and
                          (lStCycleRevision <> '');

  MemoSynthese.Visible := JaiLeRoleCompta( RcReviseur ) and
                          (VH^.Revision.Plan <> '') and
                          (lStCycleRevision <> '');
end;

////////////////////////////////////////////////////////////////////////////////

end.
