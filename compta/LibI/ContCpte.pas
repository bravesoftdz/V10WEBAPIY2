// FQ 19694 - mbo - 18.06.2007 : ajout de Ctl+N pour ouvrir l'écran de création de compte

unit ContCpte;

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
  Hctrls,
  Grids,
  ExtCtrls,
  HPanel,
  ComCtrls,
  HTB97,
  HEnt1,
  HSysMenu,
  {$IFDEF EAGLCLIENT}
  utob,
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  UiUtil,
  ImOutGen;

type
  TFContCpte = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    ToolbarButton973: TToolbarButton97;
    HPanel1: THPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    GridCompte: THGrid;
    HMTrad: THSystemMenu;
    Binsert: TToolbarButton97;
    procedure GridCompteDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GrilleKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Déclarations privées }
    fListeCompte : TList;
    procedure CreationCompte;
  public
    { Déclarations publiques }
  end;

procedure CreationComptesImmo (var L : TList);

implementation

{$IFDEF SERIE1}
uses {listcpt,}integecr, Aglinit,
{$IFDEF eAGLClient}
    MaineAGL //YCP 25/08/05
{$ELSE}
    FE_Main
{$ENDIF}
     ; //XMG 22/04/03
{$ELSE}
uses
  CPGeneraux_TOM;
{$ENDIF}

{$R *.DFM}

procedure CreationComptesImmo (var L : TList);
var
  FContCpte: TFContCpte;
begin
  FContCpte:=TFContCpte.Create(Application);
  FContCpte.fListeCompte := L;
  try
    FContCpte.ShowModal;
  finally
    L := FContCpte.fListeCompte;
    if L.Count = 0 then
    begin
      L.Free;
      L := nil;
    end;
    FContCpte.Free;
  end;
end;

procedure TFContCpte.CreationCompte;
var Q:TQuery;
    Compte : string;
    i : integer;
    ARecord : ^TDefCompte;
begin
Compte := GridCompte.Cells[0,GridCompte.Row];
if not ExisteSQL('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL = "'+Compte+'"') then
  begin
  {$IFDEF SERIE1}
  AGLLanceFiche('C','GENERAUX',Compte,'',actiontostring(taCreat)+';') ;
  {$ELSE}
    FicheGene (nil,'',Compte,taCreatOne,0) ;
  {$ENDIF}
  Q:=OpenSQL ('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL = "'+Compte+'"',true);
  if not Q.EOF then
  begin
    GridCompte.Cells[GridCompte.Col+1,GridCompte.Row] := Q.FindField('G_LIBELLE').AsString;
    for i := 0 to fListeCompte.Count - 1 do
      begin
      ARecord := fListeCompte.Items[i];
      if ARecord^.Compte = Compte then
        begin
        fListeCompte.Delete(i);
        break;
        end;
      end;
    end;
  Ferme (Q);
  end ;
end;

procedure TFContCpte.GridCompteDblClick(Sender: TObject);
begin
  CreationCompte;
end;

procedure TFContCpte.BinsertClick(Sender: TObject);
begin
  CreationCompte;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 18/06/2007
Modifié le ... :   /  /
Description .. : ajout de la fonctionnalité CTRL + N = ouvre la fenêtre de
Suite ........ : création de compte
Mots clefs ... : FQ 19694
*****************************************************************}
procedure TFContCpte.GrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //IF (Shift=[ssCtrl]) and (key = Ord('N'))  Then
  //   BinsertClick(Sender);
  if (Key = vk_Nouveau) and (ssCtrl in Shift) and (bInsert.Visible) and
    (bInsert.Enabled) then
  begin
    key := 0;
    BinsertClick(Sender);
  end;
end;

procedure TFContCpte.FormShow(Sender: TObject);
var  i,j : integer;
     ARecord : ^TDefCompte;
     bTrouve : boolean;
begin
  for i := 0 to fListeCompte.Count - 1 do
  begin
    bTrouve := false;
    ARecord := fListeCompte.Items[i];
    for j := 1 to GridCompte.RowCount -1 do
    begin
      if GridCompte.Cells[0,j] = ARecord^.Compte then
      begin
        bTrouve := true;break;
      end;
    end;
    if not bTrouve then
    begin
      GridCompte.RowCount := i+2;
      GridCompte.Cells[0,i+1] := ARecord^.Compte;
    end;
  end;
end;

procedure TFContCpte.FormClose(Sender: TObject; var Action: TCloseAction);
begin GridCompte.VidePile (False); end;

end.
