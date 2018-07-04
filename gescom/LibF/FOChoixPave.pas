{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Choix d'un élément dans une tablette par l'intermédiaire du
Suite ........ : pavé tactile.
Mots clefs ... : FO
*****************************************************************}
unit FOChoixPave;

interface

uses
  Forms, Graphics, HSysMenu, HPanel, StdCtrls, Buttons, Classes, Controls,
  ExtCtrls, Sysutils, Math, HCtrls, HEnt1, Choix, Kb_Ecran;

function FOChoisirPave(Titre, Tablette, Plus: string; Couleur: TColor = clBtnFace; Image: Integer = -1; Confirme: Boolean = False; HelpId: Integer = 0; Defaut:
  string = ''): string;

type
  TFFOChoixPave = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    BValide: TBitBtn;
    BAnnule: TBitBtn;
    BAide: TBitBtn;
    PPave: THPanel;
    procedure BAideClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    Tablette: string; // Nom de la tablette
    Plus: string; // Clause plus
    CodeChoisi: string; // Code choisi
    PnlBoutons: TClavierEcran; // pavé tactile
    procedure CreePave(Tablette, Plus: string; Couleur: TColor; Image: Integer);
    procedure CalculNbBoutons;
    procedure SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
  public
  end;

implementation

uses
  FOUtil, TickUtilFO;

{$R *.DFM}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : Choix d'un élément dans une tablette par l'intermédiaire du
Suite ........ : pavé tactile s'il existe ou grâce à la fonction Choisir.
Mots clefs ... : FO
*****************************************************************}

function FOChoisirPave(Titre, Tablette, Plus: string; Couleur: TColor = clBtnFace; Image: Integer = -1; Confirme: Boolean = False; HelpId: Integer = 0; Defaut:
  string = ''): string;
var sTable, sColonne, sSelect, sWhere, sOrdre, sPrefixe: string;
  FChoix: TFFOChoixPave;
  Ind: Integer;
begin
  Result := '';
  if FOExisteClavierEcran then
  begin
    FChoix := TFFOChoixPave.Create(Application);
    try
      FChoix.Tablette := Tablette;
      FChoix.Plus := Plus;
      FChoix.Caption := Titre;
      if Defaut <> '' then
        FChoix.Caption := FChoix.Caption + ' : ' + RechDom(Tablette, Defaut, False, Plus) + ' (' + Defaut + ')';
      FChoix.HelpContext := HelpId;
      FChoix.BValide.Visible := Confirme;
      FChoix.CodeChoisi := '';
      FChoix.CreePave(Tablette, Plus, Couleur, Image);
      if FChoix.ShowModal = mrOK then Result := FChoix.CodeChoisi;
    finally
      FChoix.Free;
    end;
  end else
  begin
    GetCorrespType(Tablette, sTable, sSelect, sWhere, sPrefixe, sColonne);
    ChangeWhereCombo(sWhere);
    ChangeWhereTT(sWhere, Plus, False);
    Ind := Pos('DISTINCT ', sSelect);
    if Ind > 0 then
    begin
      sColonne := sSelect;
      sSelect := Copy(sSelect, Ind + 9, Length(sSelect));
      sOrdre := sSelect;
    end else
    begin
      sOrdre := sColonne;
      Ind := Pos(' ORDER BY ', sWhere);
      if Ind > 0 then
      begin
        sOrdre := Copy(sWhere, Ind + 9, Length(sWhere));
        Delete(sWhere, Ind, Length(sWhere));
      end;
    end;
    Result := Choisir(Titre, sTable, sColonne, sSelect, sWhere, sOrdre, False, False, Helpid, Defaut);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CreePave : création du pavé tactile
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.CreePave(Tablette, Plus: string; Couleur: TColor; Image: Integer);
begin
  PnlBoutons := TClavierEcran.Create(Self);
  with PnlBoutons do
  begin
    Parent := PPave;
    Align := alClient;
    Height := PPave.Tag;
    ColorNb := 6;
    ChargeAut := False;
    LanceBouton := SaisieClavierEcran;
    Caisse := FOCaisseCourante;
    CalculNbBoutons;
    ClcPosition := pcLeft;
    ClcVisible := False;
    LanceChargePageMenu('123', Tablette, Plus, Couleur, Image);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculNbBoutons : calcule le nombre de boutons en hauteur et en largeur
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.CalculNbBoutons;
var Stg: string;
  NbBtn, NbLarg, NbHaut: Integer;
  ComboPage: THValComboBox;
begin
  // caractéristiques du pavé de la saisie de ticket
  Stg := FODonneParamClavierEcran(True);
  NbLarg := ValeurI(ReadTokenSt(Stg));
  NbHaut := ValeurI(ReadTokenSt(Stg));
  // calcul du nombre de valeus possibles
  ComboPage := THValComboBox.Create(Self);
  ComboPage.Vide := False;
  ComboPage.Visible := False;
  ComboPage.Parent := Self;
  ComboPage.DataType := Tablette;
  ComboPage.Plus := Plus;
  if Trim(Plus) = '' then
    ComboPage.Exhaustif := exNon
  else ComboPage.Exhaustif := exPlus;
  NbBtn := ComboPage.Items.Count;
  ComboPage.Free;
  // calcul du nombre de boutons avec comme maximum ceux définis pour la saisie de ticket
  // et comme minimum 3 en hauteur et 5 en largeur
  if NbBtn < (NbLarg * NbHaut) then
  begin
    if NbBtn <= 15 then
    begin
      NbHaut := 3;
      NbLarg := 5;
    end else
    begin
      NbHaut := Trunc(Power(NbBtn, 0.5));
      if NbHaut < 1 then NbHaut := 1;
      NbLarg := (NbBtn div NbHaut) + 1;
      if NbLarg < 1 then NbLarg := 1;
    end;
  end;
  PnlBoutons.NbrBtnWidth := NbLarg;
  PnlBoutons.NbrBtnHeight := NbHaut;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieClavierEcran : interprète les boutons du pavé
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
var Ind: Integer;
  Stg: string;
begin
  CodeChoisi := Extra;
  Stg := Caption;
  Ind := Pos(' : ', Stg);
  if Ind > 0 then Delete(Stg, Ind, Length(Stg));
  if CodeChoisi <> '' then
    Stg := Stg + ' : ' + RechDom(Tablette, CodeChoisi, False, Plus) + ' (' + CodeChoisi + ')';
  Caption := Stg;
  if not BValide.Visible then ModalResult := mrOk;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BAideClick : OnClick sur le bouton BAIDE
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.FormShow(Sender: TObject);
begin
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormDestroy :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFFOChoixPave.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

end.
