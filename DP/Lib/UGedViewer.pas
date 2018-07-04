unit UGedViewer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTob,
  HSysMenu, OleCtrls, SHDocVw_TLB, HTB97, HPanel, UIUtil, HCtrls, HMsgBox;

type
  TFGedViewer = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BModifier: TToolbarButton97;
    BChercher: TToolbarButton97;
    BApercu: TToolbarButton97;
    Web0: TWebBrowser_V1;
    HMTrad: THSystemMenu;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    { Déclarations privées }
    SDocGUID : String; // n° document réellement chargé, 0 sinon
    strUrl   :string; // url à afficher
    FTempFileName : String;
  public
    { Déclarations publiques }
    Retour : Boolean;
  end;

function ShowGedViewer (sDocGUID:string; bForceModale:boolean=False; strUrl:string=''): Boolean;

///////// IMPLEMENTATION //////////
implementation

Uses UGedFiles, UtilGed
{$IFDEF EAGLCLIENT}
  ,MenuOlx
{$ELSE}
  ,MenuOlg
{$ENDIF}
;

function ShowGedViewer (sDocGUID:string; bForceModale: Boolean=False; strUrl:string=''): Boolean;
// ouverture de la présente fiche
// qui permet de visualiser le fichier principal d'un document
// True si ouvert, False si inexistant
var  F : TFGedViewer;
     PP : THPanel;
BEGIN
  Result := False;
  if (SDocGUID='') and (strUrl='') then
     exit;

  F := TFGedViewer.Create (Application);
  F.SDocGUID := SDocGUID;
  F.strUrl   := strUrl; // $$$ JP 19/12/06

  PP := Nil;
  if Not bForceModale then
    // Ferme la précédente fiche éventuelle
    if PrepareInside then PP := FindInsidePanel;

  if PP=nil then
    begin
    try
      F.ShowModal ;
      Result := F.Retour;
    finally
      F.Free ;
      end ;
    end
  else
    begin
    InitInside(F,PP);
    F.Show;
    F.SetFocus;
    Result := F.Retour;
    end;
end;


{$R *.DFM}

procedure TFGedViewer.FormClose(Sender: TObject; var Action: TCloseAction);
var
    Flags, TargetFrameName, PostData, Headers: OleVariant;
begin
  // Visualisation page vide
  Web0.Navigate ('', Flags, TargetFrameName, PostData, Headers);
  Application.ProcessMessages;
  // Purge fichier temporaire
  if (FTempFileName<>'') and V_GedFiles.FileExists(FTempFileName) then
    DeleteFile(FTempFileName);
end;

procedure TFGedViewer.FormShow(Sender: TObject);
var {Url,} Titre : String ;
    Flags, TargetFrameName, PostData, Headers: OleVariant;
begin
  Retour := False;

  // Recherche du fichier principal du doc, et propose un titre
  if strUrl = '' then
  begin
       FTempFileName := ExtraitDocument(SDocGUID, Titre);
       if FTempFileName='' then
          exit;
  end
  else
      Titre := strUrl;

  Self.Caption := Titre;
  UpdateCaption (Self);

  // Visualisation du fichier extrait ou de l'url déjà spécifié
  if strUrl = '' then
  begin
       if V_GedFiles.FileExists(FTempFileName) then
           strUrl := 'file:///' + ExpandFileName(FTempFileName)
       else
           strUrl := 'about:blank';
  end;

  // Affichage
  try
    Web0.Navigate (strUrl, Flags, TargetFrameName, PostData, Headers);
  except
    On E:Exception do PGIInfo(E.Message, Self.Caption);
  end;

  Retour := True;
end;


end.

