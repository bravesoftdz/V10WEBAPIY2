{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /    
Description .. : Détope du fichier TOX
Mots clefs ... : TOX;DETOPAGE
*****************************************************************}
unit ToxDetop;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ComCtrls, StdCtrls, ExtCtrls, HPanel, uTob, hmsgbox, Grids, Hctrls,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} ToxMoteur, HEnt1, UIUtil, HSysMenu ;



type
  TDetopageTox = class(TForm)
    HPanel1: THPanel;
    HPanel3: THPanel;
    OD: TOpenDialog;
    CompteRendu: TMemo;
    PMain: THPanel;
    Label3: TLabel;
    Bopen: TToolbarButton97;
    HMTrad: THSystemMenu;
    Dock: TDock97;
    Divers: TToolbar97;
    baide: TToolbarButton97;
    procedure FormShow(Sender: TObject) ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure BopenClick(Sender: TObject);
    procedure BcancelClick(Sender: TObject);
    function  ChargeTOX (TT : TOB) : Integer ;
    procedure FormCreate(Sender: TObject);
    procedure baideClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    TobADetoper : TOB ;
  end;

////////////////////////////////////////////////////////////////////////////////
// Procédures ou fonctions à exporter
////////////////////////////////////////////////////////////////////////////////
  procedure DetopeTox ;

implementation

//uses AglToxRun;

{$R *.DFM}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Création de la forme pour détoper un fichier TOX
Mots clefs ... : TOX;DETOPAGE
*****************************************************************}
procedure DetopeTox ( ) ;
var
  X  : TDetopageTox;
  PP : THPanel ;
begin
  PP:=FindInsidePanel ;
  X := TDetopageTox.Create ( Application ) ;
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
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /    
Description .. : Charge la TOX du fichier dans la TOB "TobADetoper"
Mots clefs ... : TOX
*****************************************************************}
function TDetopageTox.ChargeTOX(TT : TOB) : Integer ;
BEGIN
  TobADetoper:=TT ;
  result:=0 ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /    
Description .. : TobEDetoper : Libération de la TOB de récupération
Mots clefs ... : TOX
*****************************************************************}
procedure TDetopageTox.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree ;
end;

procedure TDetopageTox.FormDestroy(Sender: TObject);
begin
  if TobADetoper <> nil then TobADetoper.Free ;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /    
Description .. : Ouverture du fichier et détopage
Mots clefs ... : TOX
*****************************************************************}
procedure TDetopageTox.BopenClick(Sender: TObject);
var EtatTox : integer ;
begin
  if OD.Execute then
  begin
    CompteRendu.lines.add ('   ' + TraduireMemoire ('Détopage du fichier') + ' ' + OD.FileName) ;
    //
    // Chargement du fichier en TOB
    //
    TobADetoper := Tob.Create ( 'TOB MAITRE', Nil, -1 ) ;
    TOBLoadFromBinFile (OD.FileName, ChargeTOX, nil);
    //
    // intégration de la TOB
    //
    EtatTox := DetopageDeLaTox (TobADetoper, CompteRendu) ;
    //
    // Ré-écriture de la TOB dans le fichier
    //
    if EtatTox = ToxIntegrer then TobADetoper.SaveToBinFile (OD.FileName, False, True, True, False);
    //
    // Libération de la TOB
    //
    TobADetoper.free;
    TobADetoper:=nil;
  end ;
end;

procedure TDetopageTox.BcancelClick(Sender: TObject);
begin
  Close ;
end;

procedure TDetopageTox.FormCreate(Sender: TObject);
begin
  OD.Filter := TraduireMemoire ('Fichiers des échanges PGI') + ' (*.TOX)|*.TOX' ;
end;

procedure TDetopageTox.FormShow(Sender: TObject);
begin
// Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self) ;
end;

procedure TDetopageTox.baideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

end.
