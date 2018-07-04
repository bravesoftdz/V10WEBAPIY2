{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/09/2001
Modifié le ... : 24/09/2001
Description .. : Affiche un message pour avertir l'utilisateur que la
Suite ........ : caisse est en attente de communication avec le site central.
Mots clefs ... : FO
*****************************************************************}
unit FOAttenteComm;

interface

uses
  Forms, Vierge, Hctrls, StdCtrls, HFLabel, HSysMenu, Controls, HTB97, Classes,
  Windows,
  {$IFDEF TOXCLIENT}
  UTox,
  {$ENDIF}
  HPanel, UIUtil, HEnt1;

type
  TFFOAttenteComm = class(TFVierge)
    LAttention: TFlashingLabel;
    LTexte: THLabel;
    LTexte1: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure FOMiseAttenteComm(Titre: string);

implementation

{$R *.DFM}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/09/2001
Modifié le ... : 24/09/2001
Description .. : Affiche un message pour avertir l'utilisateur que la
Suite ........ : caisse est en attente de communication avec le site central.
Mots clefs ... : FO
*****************************************************************}

procedure FOMiseAttenteComm(Titre: string);
var XX: TFFOAttenteComm;
  PP: THPanel;
begin
  SourisSablier;
  PP := FindInsidePanel;
  XX := TFFOAttenteComm.Create(Application);
  if Titre <> '' then XX.Caption := Titre;
  if PP = nil then
  begin
    try
      XX.ShowModal;
    finally
      XX.Free;
    end;
    SourisNormale;
  end else
  begin
    InitInside(XX, PP);
    XX.Show;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/09/2001
Modifié le ... : 24/09/2001
Description .. : FormShow
Mots clefs ... :
*****************************************************************}

procedure TFFOAttenteComm.FormShow(Sender: TObject);
const MARGE = 20;
begin
  inherited;
  LAttention.Width := Width - (2 * MARGE);
  LAttention.Left := Left + MARGE;
  LTexte.Width := Width - (2 * MARGE);
  LTexte.Left := Left + MARGE;
  LTexte1.Width := Width - 40;
  LTexte1.Left := Left + MARGE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 16/10/2003
Modifié le ... : 16/10/2003
Description .. : FormKeyDown
Mots clefs ... :
*****************************************************************}

procedure TFFOAttenteComm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
{$IFDEF TOXCLIENT}
var
  WhenStart: TDateTime;
{$ENDIF}
begin
  inherited;
  {$IFDEF TOXCLIENT}
  if Key = VK_ESCAPE then
  begin
    // pas de sortie par la touche Escape possible si une communication est en cours
    if AglStatusTox(WhenStart) then Key := 0;
  end;
  {$ENDIF}
end;

end.
