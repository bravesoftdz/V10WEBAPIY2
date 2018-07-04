{***********UNITE*************************************************
Auteur  ...... : Guillaume Fontana
Créé le ...... : 20/09/2000
Modifié le ... : 20/09/2000
Description .. : Choix d'une icone parmi la liste des icones PGI
Suite ........ : Renvoie l'icone sélectionné.
Mots clefs ... : ICONE
*****************************************************************}
unit HRListeIcone;

interface

uses
  Windows,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  HTB97,
  ExtCtrls,
  ComCtrls,
  ImgList,
  hent1,
  M3FP,
  HDrawXP,
  HSysMenu, HImgList;

  function ListeIcone(NoIconePossible: Boolean): integer;

  procedure ChargeIcone(NumIcone: Integer; Image: TImage);

type
  TFicheIcone = class(TForm)
    LIcone: TListView;
    LImage: THImageList;
    Dock971: TDock97;
    Toolbar971: TToolbar97;
    BAnnul: TToolbarButton97;
    Bok: TToolbarButton97;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BOkClick(Sender: TObject);
    procedure BAnnuClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    Ok: Boolean;
    PasIcone: Boolean;
    Icon: TIcon;
  public
    { Déclarations publiques }
    NbImage: Integer;
    IconeSelection: Integer;
  end;
var
  FicheIcone: TFicheIcone;

implementation

{$R *.DFM}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : GF
Créé le ...... : 20/09/2000
Modifié le ... : 20/09/2000
Description .. : Affiche la liste des icones PGI, et retourne le numéro d'icone
Suite ........ : selectionné
Suite ........ : - Paramètre NoIconePossible : Aucune selection possible
Suite ........ : (renvoie alors 9999)
Mots clefs ... : ICONE
*****************************************************************}
function ListeIcone(NoIconePossible: Boolean): Integer;
begin
  with TFicheIcone.Create(Application) do
  begin
    PasIcone := NoIconePossible;
    ShowModal;
    Result := IconeSelection;
    Free;
  end;
end;

procedure TFicheIcone.FormShow(Sender: TObject);
var
  i,c: integer;
  Executer: Boolean;
  Tlvi: tlistitem;
begin
  // Chargement des images
  Ok := False;
  Executer := True;
  i := 0;
  c := 0;
  while Executer do
  begin
    icon := TIcon.Create;
    if V_PGI.Draw2003 then
     // LoadZipIconFromRessources('N' + IntToStr(i),32,Icon)
      Icon.Handle := LoadIcon(AglHInstance,PChar('N' + IntToStr(i)))
    else
      Icon.Handle := LoadIcon(AglHInstance,PChar('Z' + IntToStr(i)));

    if icon.Handle = 0 then
    begin
      break;
    end
    else
      LImage.AddIcon(ICON);
    // Création des images
    Tlvi := LIcone.Items.Add;
    Tlvi.imageindex := i;
    Tlvi.Caption := IntToStr(i);
    inc(c);
    if c = 3 then
    begin
      c := 0;
    end;
    Icon.Free;
    inc(i);
  end;
  NbImage := i;
  if PasIcone = True then
  begin
    Tlvi := LIcone.Items.Insert(0);
    Tlvi.imageindex := 9999;
    Tlvi.Caption := 'Aucun Icône';
  end;
end;

procedure TFicheIcone.BOkClick(Sender: TObject);
begin

  if LIcone.SelCount > 0 then
  begin
    IconeSelection := LIcone.Selected.Index;
    {cas NoIconePossible = True}
    if (PasIcone = true) then
    begin
      if IconeSelection = 0 then
        IconeSelection := 9999
      else
        IconeSelection := LIcone.Selected.Index - 1;
    end;
  end
  else
    IconeSelection := -1;
  ok := true;
  close;
end;

procedure TFicheIcone.BAnnuClick(Sender: TObject);
begin
  IconeSelection := -1; // Certainemement pas utile à voir...
  Close;
end;

procedure TFicheIcone.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  // Si aucune selection et fermeture
  if Ok = False then  IconeSelection := -1;

  icon.free;

end;

// Charge un icone AGL dans un TImage
procedure ChargeIcone(NumIcone: Integer; Image: TImage);
var
  Icons: TIcon;
begin
  if Image <> nil then
  begin
    // Creation de l'icone
    Icons := TIcon.Create;
    // Charge l'icone de numéro "NumIcone"
    // Look 2003 on charge les iconnes zippés
    if V_PGI.Draw2003 then
      //LoadZipIconFromRessources('Z' + IntToStr(NumIcone),32,Icons) //permet de lire la new icon dans la ressource avec la taille spécifiée
      Icons.Handle := LoadIcon(AglHInstance,PChar('N' + IntToStr(NumIcone)))
    else
      Icons.Handle := LoadIcon(AglHInstance,PChar('Z' + IntToStr(NumIcone)));
    if icons.Handle <> 0 then
    begin
      // Affectation à l'objet TImage
      Image.Picture.Icon := Icons;
    end;
    Icons.Free;
  end;
end;

initialization

end.
