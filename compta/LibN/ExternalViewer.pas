unit ExternalViewer;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, Menus, ExtCtrls, HPanel, Hctrls,uTob,uLibSaisieDoc;

type
  TViewerExt = class(TForm)
    PBouton       : TToolWindow97;
    BValider      : TToolbarButton97;
    BFerme        : TToolbarButton97;
    HelpBtn       : TToolbarButton97;
    BMENUVIEWER   : TToolbarButton97;
    POPUPVIEWER: TPopupMenu;
    VIEWERNEXT: TMenuItem;
    VIEWERPREC: TMenuItem;
    VIEWERMOVE: TMenuItem;
    VIEWERMOVEACH: TMenuItem;
    VIEWERMOVEBQE: TMenuItem;
    VIEWERMOVECAI: TMenuItem;
    VIEWERMOVEOD: TMenuItem;
    VIEWERMOVEVTE: TMenuItem;
    VIEWERMOVECORBEILLE: TMenuItem;
    VIEWERPRINT: TMenuItem;
    VIEWERPAN: TMenuItem;
    VIEWERZOOMAV: TMenuItem;
    VIEWERZOOMAR: TMenuItem;
    VIEWERFITPAGE: TMenuItem;
    VIEWERFITLARGE: TMenuItem;
    VIEWERFIT100: TMenuItem;
    VIEWERROTATE: TMenuItem;
    PBOTTOM: THPanel;
    SPBOTTOM: THSplitter;
    SPRIGHT: THSplitter;
    SPTOP: THSplitter;
    SPLEFT: THSplitter;
    PTOP: THPanel;
    PLEFT: THPanel;
    PRIGHT: THPanel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    FTSDoc        : TSaisieDoc ;
    FTSDocInt     : TSaisieDoc ;
    FTSP          : TObject ;
    procedure EchangeViewer ( Src : TSaisieDoc ; Dst : TSaisieDoc );
    procedure SaveAffichage;
    procedure LoadAffichage;
    procedure Close;
  public
    { Déclarations publiques }
    property TSDocExt  : TSaisieDoc   read FTSDoc     write FTSDoc;
    property TSDocInt  : TSaisieDoc   read FTSDocInt  write FTSDocInt;
    property TSP       : TObject      read FTSP       write FTSP;
    procedure InitViewer;
  end;

var
  ViewerExt: TViewerExt;

implementation
  uses  hmsgbox,hent1,uLibSaisiePiece;
{$R *.dfm}

procedure TViewerExt.InitViewer;
begin
  FTSDoc.InitViewer( PBOTTOM, PRIGHT, PLEFT, PTOP, SPBOTTOM, SPRIGHT, SPLEFT, SPTOP ) ;
  FTSDoc.SetMenuPop( POPUPVIEWER, BMENUVIEWER, nil) ;  
  POPUPVIEWER.OnPopup := FTSDoc.Onpopupviewer ;
  FTSDoc.InitPopupViewer(False);
end;

procedure TViewerExt.FormShow(Sender: TObject);
begin
  if not(assigned(TSaisiePiece(FTSP))) then Exit ;
  FTSDocInt         := TSaisiePiece(FTSP).TSDocInt;
  FTSDoc            := TSaisiePiece(FTSP).TSDocExt;
  TSaisiePiece(FTSP).Viewer := 'EXT';
  // On désactive si besoin l'affichage du viewer interne
  if FTSDocInt.IsVisible then
  begin
     FTSDocInt.MontreViewer(false);
     TSaisiePiece(FTSP).ResizeGrille ;
  end;
  LoadAffichage;
  InitViewer;                                        
  EchangeViewer(TSaisiePiece(FTSP).TSDocInt,FTSDoc);
  FTSDoc.ActiverViewer(TSaisiePiece(FTSP).GetTobMasqueParam, true);
  FTSDoc.ChargeDocument(false);
  FTSDocInt.SetVisibleBDoc(true);
  FTSDoc.PanelActif.Align := alclient;
end;

procedure TViewerExt.EchangeViewer ( Src : TSaisieDoc ; Dst : TSaisieDoc );
begin
  Dst.Fichier := Src.Fichier;
  Dst.SearchRec := Src.SearchRec ;
end;

procedure TViewerExt.BValiderClick(Sender: TObject);
begin
  PostMessage (Self.Handle, WM_CLOSE, 0, 0);
end;


procedure TViewerExt.SaveAffichage;
begin
  SaveSynRegKey('VisuHeight'  ,Self.Height  ,true);
  SaveSynRegKey('VisuLeft'    ,Self.Left    ,true);
  SaveSynRegKey('VisuTop'     ,Self.Top     ,true);
  SaveSynRegKey('VisuWidth'   ,Self.Width   ,true);
end;

procedure TViewerExt.LoadAffichage;
begin
  Self.Height := GetSynRegKey('VisuHeight'  ,Self.Height  ,true);
  Self.Left   := GetSynRegKey('VisuLeft'    ,Self.Left    ,true);
  Self.Top    := GetSynRegKey('VisuTop'     ,Self.Top     ,true);
  Self.Width  := GetSynRegKey('VisuWidth'   ,Self.Width   ,true);
end;

procedure TViewerExt.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
   canClose := true;
end;

procedure TViewerExt.FormDestroy(Sender: TObject);
begin
  Close;
end;

procedure TViewerExt.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Close;
  Action := caFree;
end;

procedure TViewerExt.Close;
begin
  if assigned(FTSDoc) then
  begin                    
     SaveAffichage;
     if assigned(FTSDocInt) then
     begin
        FTSDocInt.SetVisibleBDoc(false);
        TSaisiePiece(FTSP).Viewer := 'INT';
        SaveSynRegKey('PosVisu'  ,'INT'  ,true);
        EchangeViewer( FTSDoc , FTSDocInt);
        // On réactive si besoin l'affichage du viewer interne
        TSaisiePiece(FTSP).SetDocActif(FTSDocInt);
        // MAJ grille suivant paramètrage
        TSaisiePiece(FTSP).UpdateMasqueSaisie ;
        if FTSDocInt.IsActif then
        begin
           FTSDocInt.MontreViewer(true);  
           TSaisiePiece(FTSP).paramViewer ;
        end;
     end;
     FTSDoc.DesactiveViewer ;
     FTSDoc.DesInitViewer ;
     FTSDoc := nil;  
     FTSDocInt := nil;
  end;
  TSaisiePiece(FTSP).ViewerExt := nil;
end;

procedure TViewerExt.FormResize(Sender: TObject);
begin
  if assigned(FTSDoc) and assigned(FTSDoc.InternalViewer) then FTSDoc.ViewerFitLargeClick(Sender);
end;

procedure TViewerExt.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case Key Of
      VK_ESCAPE :
      begin
         PostMessage (Self.Handle, WM_CLOSE, 0, 0);
         Key := 0;
      end ;
      else
      begin
         if assigned(TSP) and assigned(TSaisiePiece(FTSP).TSDoc) then
            TSaisiePiece(FTSP).ViewerKeyDown(Sender,Key,Shift);
         Key := 0;
      end;
   end;
end;

end.
