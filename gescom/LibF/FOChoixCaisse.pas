{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 18/05/2001
Description .. : Choix de la caisse de travail
Mots clefs ... : FO
*****************************************************************}
unit FOChoixCaisse;

interface

uses
  Forms, HSysMenu, StdCtrls, Hctrls, HTB97, Controls, Classes, ExtCtrls,
  {$IFNDEF EAGLCLIENT}
  dbtables,
  {$ENDIF}
  HMsgBox, HEnt1, HPanel, UTOB;

function FOChoixCodeCaisse: Boolean;

type
  TFChoixCaisse = class(TForm)
    PMAIN: THPanel;
    TGPK_ETABLISSEMENT: THLabel;
    bConnect: TToolbarButton97;
    bQuit: TToolbarButton97;
    HMTrad: THSystemMenu;
    TGPK_CAISSE: THLabel;
    GPK_CAISSE: THValComboBox;
    TGPK_DEPOT: THLabel;
    GPK_DEPOT: THValComboBox;
    GPK_ETABLISSEMENT: THValComboBox;
    LGPK_ETABLISSEMENT: TEdit;
    LGPK_DEPOT: TEdit;
    procedure FormShow(Sender: TObject);
    procedure bConnectClick(Sender: TObject);
    procedure bQuitClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure GPK_CAISSEChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private { Déclarations privées }
    TOBCaisse: TOB; // TOB de la caisse choisie
    Connexion: Integer; // Indicateur de connexion établie
    procedure ChargeCaisse;
  public { Déclarations publiques }
  end;

implementation

uses
  FOUtil, FODefi;

{$R *.DFM}

///////////////////////////////////////////////////////////////////////////////////////
//  ValideChoixCaisse : valide la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

function ValideChoixCaisse(TOBCaisse: TOB): Boolean;
var sTitre: string;
begin
  Result := False;
  if TOBCaisse = nil then Exit;
  if (TOBCaisse.GetValue('GPK_CAISSE') <> '') and (TOBCaisse.GetValue('GPK_ETABLISSEMENT') <> '') and
    (TOBCaisse.GetValue('GPK_DEPOT') <> '') then
  begin
    sTitre := TraduireMemoire('Caisse') + ' : ' + TOBCaisse.GetValue('GPK_CAISSE')
      + ' ' + TOBCaisse.GetValue('GPK_LIBELLE');
    // Verifie si la caisse est déjà verouillée.
    if ((FOOptionConnexionCaisse(cfoVerifVerrou)) and (TOBCaisse.GetValue('GPK_INUSE') = 'X') and
       (not FOCaisseIsMonoPoste)) then
    begin
      PGIError('La caisse est déjà en cours d''utilisation !', sTitre);
    end else
    begin
      // Verrouille la caisse
      if ((not FOLockCaisse(TOBCaisse.GetValue('GPK_CAISSE'))) and
         (not FOCaisseIsMonoPoste)) then
      begin
        PGIError('La caisse n''a pas pu être verrouillée !', sTitre);
      end else
      begin
        // Mémorisation du paramètrage de la caisse choisie dans la variable globale PGI
        FOChargeVHGCCaisse(TOBCaisse.GetValue('GPK_CAISSE'));
        // Enregistrement dans la registry du nom de la caisse utilisée
        /////SaveSynRegKey(CAISSEPARDEFAUT, TOBCaisse.GetValue('GPK_CAISSE'), True) ;
        Result := True;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  UneSeuleCaisse : si une seule caisse est définie, pas de besoin de choix
///////////////////////////////////////////////////////////////////////////////////////

function UneSeuleCaisse: Boolean;
var
  QQ: TQuery;
  TOBC: TOB;
  sSelect: string;
begin
  Result := False;
  sSelect := 'SELECT COUNT(*) FROM PARCAISSE WHERE GPK_FERME="-"';
  QQ := OpenSQL(sSelect, True);
  if not (QQ.EOF) and (QQ.Fields[0].AsInteger = 1) then
  begin
    Ferme(QQ);
    sSelect := 'SELECT * FROM PARCAISSE WHERE GPK_FERME="-"';
    QQ := OpenSQL(sSelect, True);
    if not QQ.EOF then
    begin
      TOBC := TOB.Create('PARCAISSE', nil, -1);
      TOBC.SelectDB('', QQ, False);
      Result := ValideChoixCaisse(TOBC);
      TOBC.Free;
    end;
 end;
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CaissePoste : récupére la caisse attachée au poste de traval
///////////////////////////////////////////////////////////////////////////////////////

function CaissePoste(var ChoixCaisseOk: Boolean): Boolean;
var sCaisse, sSql, sTexte, sTitre: string;
  QQ: TQuery;
  TOBC: TOB;
begin
  // Lecture de la dernière caisse utilisée dans la registry
  ChoixCaisseOk := False;
  sCaisse := GetSynRegKey(CAISSEPARDEFAUT, '', True);
  if sCaisse <> '' then
  begin
    Result := True;
    sTitre := TraduireMemoire('Caisse') + ' : ' + sCaisse;
    sTexte := TraduireMemoire('La caisse attachée à la station');
    sSql := 'SELECT * FROM PARCAISSE WHERE GPK_CAISSE="' + sCaisse + '"';
    QQ := OpenSQL(sSql, True);
    if QQ.EOF then
    begin
      sTexte := sTexte + TraduireMemoire(' n''existe pas.');
      PGIBox(sTexte, sTitre);
    end else
      if QQ.FindField('GPK_FERME').AsString <> '-' then
    begin
      sTitre := sTitre + ' ' + QQ.FindField('GPK_LIBELLE').AsString;
      sTexte := sTexte + TraduireMemoire(' est fermée.');
      PGIBox(sTexte, sTitre);
    end else
    begin
      TOBC := TOB.Create('PARCAISSE', nil, -1);
      TOBC.SelectDB('', QQ, False);
      ChoixCaisseOk := ValideChoixCaisse(TOBC);
      TOBC.Free;
    end;
    Ferme(QQ);
  end else
  begin
    Result := UneSeuleCaisse;
    ChoixCaisseOk := Result;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/12/2000
Modifié le ... : 23/07/2001
Description .. : Choix de la caisse de travail
Mots clefs ... : FO
*****************************************************************}

function FOChoixCodeCaisse: Boolean;
var XX: TFChoixCaisse;
begin
  Result := False;
  if not CaissePoste(Result) then
  begin
    // écran de choix de la caisse
    XX := TFChoixCaisse.Create(Application);
    try
      XX.ShowModal;
      Result := (XX.Connexion = mrOk);
    finally
      XX.Free;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeCaisse : charge la TOB de la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.ChargeCaisse;
begin
  // Lecture du paramétrage de la caisse
  TOBCaisse.SelectDB('"' + GPK_CAISSE.Value + '"', nil);
  // Affichage des données complémentaires
  GPK_ETABLISSEMENT.Value := TOBCaisse.GetValue('GPK_ETABLISSEMENT');
  LGPK_ETABLISSEMENT.Text := GPK_ETABLISSEMENT.Text;
  GPK_DEPOT.Value := TOBCaisse.GetValue('GPK_DEPOT');
  LGPK_DEPOT.Text := GPK_DEPOT.Text;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormKeyPress : gestion du clavier
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then bConnectClick(nil); // touche Entrée
  if Key = #27 then bQuitClick(nil); // touche Echap.
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.FormShow(Sender: TObject);
var Ind: integer;
  CashDef: string;
begin
  // Lecture de la dernière caisse utilisée dans la registry
  // si aucune valeur on propose la première ligne de la liste.
  CashDef := GetSynRegKey(CAISSEPARDEFAUT, '', True);
  Ind := GPK_CAISSE.Values.IndexOf(CashDef);
  if Ind >= 0 then GPK_CAISSE.ItemIndex := Ind else GPK_CAISSE.ItemIndex := 0;
  GPK_Caisse.OnChange(nil);
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.FormCreate(Sender: TObject);
begin
  Connexion := mrNone;
  // Creation de la TOB de la caisse
  TOBCaisse := TOB.Create('PARCAISSE', nil, -1);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormDestroy
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // Libération de la TOB de la caisse
  TOBCaisse.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPK_CAISSEChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.GPK_CAISSEChange(Sender: TObject);
begin
  // Recherche de la caisse choisie
  ChargeCaisse;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bConnectClick : bouton Connexion
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.bConnectClick(Sender: TObject);
begin
  bConnect.Enabled := False;
  if ValideChoixCaisse(TOBCaisse) then Connexion := mrOk;
  bConnect.Enabled := True;
  Close;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bQuitClick : bouton Quitter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFChoixCaisse.bQuitClick(Sender: TObject);
begin
  Connexion := mrCancel;
  Close;
end;

end.
