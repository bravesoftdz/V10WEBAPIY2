unit AssistFinInv;

interface

uses
  {$IFNDEF EAGLCLIENT}
  DBTables,
  {$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  HPanel, HEnt1, UTob, EntGC;

procedure Assist_FinInvent;

type
  TFAssistFinInv = class(TFAssist)
    Panel1: TPanel;
    PTITRE: THPanel;
    Bevel1: TBevel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    GBCRITERE: TGroupBox;
    TETABLISS: THLabel;
    CB_SUPPLISTE: TCheckBox;
    CB_INVTRANS: TCheckBox;
    ETABLISSEMENTS: THMultiValComboBox;
    RG_TYPEDOC: THRadioGroup;
    TabSheet1: TTabSheet;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bFinClick(Sender: TObject);
    procedure TraitFinInvent();
    procedure bAideClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

const
  // libellés des messages
  TexteMessage: array[1..7] of string = (
    {1}'Confirmation'
    {2}, 'Les listes d''inventaire vont être supprimées pour tous les établissements sélectionnés !'
    {3}, 'Les inventaires transmis vont être supprimés pour tous les établissements sélectionnés !'
    {4}, 'Les listes et inventaires transmis vont être supprimés pour tous les établissements sélectionnés !'
    {5}, 'Le traitement est terminé.'
    {6}, 'Vous devez cocher au moins une case'
    {7}, 'Vous devez sélectionner au moins un établissement'
    );
var
  FAssistFinInv: TFAssistFinInv;

implementation

{$IFNDEF EAGLCLIENT}
uses MenuOLG;
{$ENDIF}

{$R *.DFM}

procedure Assist_FinInvent;
begin
  FAssistFinInv := TFAssistFinInv.Create(Application);
  try
    FAssistFinInv.ShowModal;
  finally
    FAssistFinInv.free;
  end;
end;

procedure TFAssistFinInv.FormShow(Sender: TObject);
begin
  inherited;
  bPrecedent.Visible := False;
  bSuivant.Visible := False;
  lAide.Caption := '';
  AglEmpileFiche(Self);

  if (ctxMode in V_PGI.PGIContexte) then ETABLISSEMENTS.Plus := 'GDE_SURSITE="X"';
end;

procedure TFAssistFinInv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  {$IFNDEF EAGLCLIENT}
  FMenuG.VireInside(nil);
  {$ENDIF}
  AglDepileFiche;
end;

procedure TFAssistFinInv.bFinClick(Sender: TObject);
begin
  inherited;
  if (CB_SUPPLISTE.Checked = False) and (CB_INVTRANS.Checked = False) then
  begin
    PGIInfo(TraduireMemoire(TexteMessage[6]), Caption);
    exit;
  end;
  if (ETABLISSEMENTS.Text = '') then
  begin
    PGIInfo(TraduireMemoire(TexteMessage[7]), Caption);
    exit;
  end;
  if (CB_INVTRANS.Checked = False) then
    if PGIAsk(TraduireMemoire(TexteMessage[2]), TraduireMemoire(TexteMessage[1])) <> mrYes then exit;
  if (CB_SUPPLISTE.Checked = False) then
    if PGIAsk(TraduireMemoire(TexteMessage[3]), TraduireMemoire(TexteMessage[1])) <> mrYes then exit;
  if (CB_SUPPLISTE.Checked = True) and (CB_INVTRANS.Checked = True) then
    if PGIAsk(TraduireMemoire(TexteMessage[4]), TraduireMemoire(TexteMessage[1])) <> mrYes then exit;
  TraitFinInvent();
  Close;
end;

// Traitement de fin d'inventaire

procedure TFAssistFinInv.TraitFinInvent();
var i: integer;
  NbListe, NbInvTrans: integer;
  MultiValCombo: string;
  CodeEtab, ListeEtab: string;
  CodeListe, CodeTrans: string;
  SQL: string;
  Q: TQuery;
  TobTemp: TOB;
begin
  // Recherche de la liste des établissements sélectionnés
  ListeEtab := '';
  MultiValCombo := ETABLISSEMENTS.Text;
  if MultiValCombo = TraduireMemoire('<<Tous>>') then
  begin
    for i := 0 to ETABLISSEMENTS.Items.Count - 1 do
    begin
      if ListeEtab <> '' then ListeEtab := ListeEtab + ',';
      ListeEtab := ListeEtab + '"' + ETABLISSEMENTS.Values[i] + '"';
    end;
  end else
  begin
    while MultiValCombo <> '' do
    begin
      CodeEtab := ReadTokenSt(MultiValCombo);
      if ListeEtab <> '' then ListeEtab := ListeEtab + ',';
      ListeEtab := ListeEtab + '"' + CodeEtab + '"';
    end;
  end;

  if ListeEtab <> '' then
  begin
    NbListe := 0;
    NbInvTrans := 0;
    if (CB_SUPPLISTE.Checked = True) then
    begin
      // ------------------------------------
      // Suppression des listes d'inventaires
      // ------------------------------------

      // Chargement en TOB, des entêtes de liste répondant aux critères de sélection suivants :
      // - Etablissement appartenant à "ListeEtab"
      // - GIE_VALIDATION="X" si uniquement les listes validées
      // - GIE_VALIDATION="-" si uniquement les listes non validées.
      TobTemp := TOB.Create('', nil, -1);
      SQL := 'SELECT GIE_DEPOT, GIE_CODELISTE FROM LISTEINVENT ' +
        'WHERE GIE_DEPOT IN (' + ListeEtab + ')';
      if RG_TYPEDOC.ItemIndex = 1 then SQL := SQL + ' and GIE_VALIDATION="X"';
      if RG_TYPEDOC.ItemIndex = 2 then SQL := SQL + ' and GIE_VALIDATION="-"';
      Q := OpenSQL(SQL, true);
      if not Q.EOF then TobTemp.LoadDetailDB('LISTEINVENT', '', '', Q, False, False);
      Ferme(Q);

      if TobTemp.Detail.Count > 0 then
      begin
        for i := 0 to TobTemp.Detail.Count - 1 do
        begin
          CodeListe := TobTemp.Detail[i].GetValue('GIE_CODELISTE');
          if (GetInfoParPiece('INV','GPP_LOT')='X') then
          begin
            SQL := 'delete from LISTEINVLOT where GLI_CODELISTE="' + CodeListe + '"';
            ExecuteSQL(SQL);
          end;
          SQL := 'delete from LISTEINVENT where GIE_CODELISTE="' + CodeListe + '"';
          ExecuteSQL(SQL);

          SQL := 'delete from LISTEINVLIG where GIL_CODELISTE="' + CodeListe + '"';
          ExecuteSQL(SQL);
          inc(NbListe);
        end;
      end;
      TobTemp.Free;
    end;

    if (CB_INVTRANS.Checked = True) then
    begin
      // ------------------------------------
      // Suppression des inventaires transmis
      // ------------------------------------

      // Chargement en TOB, des entêtes d'inventaires transmis répondant aux critères
      // de sélection suivants :
      // - Etablissement appartenant à "ListeEtab"
      // - GIT_INTEGRATION="X" si uniquement les inventaires transmis intégrés
      // - GIT_INTEGRATION="-" si uniquement les inventaires transmis non intégrés.
      TobTemp := TOB.Create('', nil, -1);
      SQL := 'SELECT GIT_DEPOT, GIT_CODETRANS FROM TRANSINVENT ' +
        'WHERE GIT_DEPOT IN (' + ListeEtab + ')';
      if RG_TYPEDOC.ItemIndex = 1 then SQL := SQL + ' and GIT_INTEGRATION="X"';
      if RG_TYPEDOC.ItemIndex = 2 then SQL := SQL + ' and GIT_INTEGRATION="-"';
      Q := OpenSQL(SQL, true);
      if not Q.EOF then TobTemp.LoadDetailDB('TRANSINVENT', '', '', Q, False, False);
      Ferme(Q);

      if TobTemp.Detail.Count > 0 then
      begin
        for i := 0 to TobTemp.Detail.Count - 1 do
        begin
          CodeTrans := TobTemp.Detail[i].GetValue('GIT_CODETRANS');
          CodeEtab := TobTemp.Detail[i].GetValue('GIT_DEPOT');
          SQL := 'delete from TRANSINVENT where GIT_CODETRANS="' + CodeTrans + '" and ' +
            'GIT_DEPOT="' + CodeEtab + '"';
          ExecuteSQL(SQL);

          SQL := 'delete from TRANSINVLIG where GIN_CODETRANS="' + CodeTrans + '" and ' +
            'GIN_DEPOT="' + CodeEtab + '"';
          ExecuteSQL(SQL);
          Inc(NbInvTrans);
        end;
      end;
      TobTemp.Free;
    end;

    //PGIInfo(TraduireMemoire(TexteMessage[5]),Caption);
    PGIInfo(IntToStr(NbListe) + ' liste(s) d''inventaire suprimée(s),'#13 + IntToStr(NbInvTrans) + ' inventaire(s) transmis supprimé(s)', Caption);
  end;
end;

procedure TFAssistFinInv.bAideClick(Sender: TObject);
begin
  inherited;
  CallHelpTopic(Self);
end;

end.
