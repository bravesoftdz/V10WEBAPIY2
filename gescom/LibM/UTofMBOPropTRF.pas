unit UTofMBOPropTRF;

interface
uses M3FP, StdCtrls, Controls, Classes, Dialogs, Grids, spin,
  HCtrls, HEnt1, HMsgBox, Hdimension, UTOB, UTOF, AGLInit, EntGC, PropoAffectTrf,
  {$IFDEF EAGLCLIENT}
  emul, Maineagl,
  {$ELSE}
  mul, DBTables, DBGrids, DB, FE_Main,
  {$ENDIF}
  Forms, SysUtils, ComCtrls, Vierge, Math, FOWaitFor;

const MaxDimChamp = 10;

type
  TOF_MBOPROPTRF = class(TOF)
  private
    procedure G_AFFCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_AFFCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_AFFRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_AFFRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure Calculcoeff;
  public
    initChamp: array[1..MaxDimChamp] of string;
    TobListe, TobAffiche, TobMethode, TobAffectationTrf: TOB;
    GLISTE, GAFFICHE, GMETHODE: THGrid;
    procedure OnClose; override;
    procedure OnLoad; override;
    procedure OnArgument(St: string); override;
    procedure OnUpdate; override;
    procedure OnMajListeBtq;
    procedure Changetaillecol;
    procedure ChangeSelectBtq(Tobind: integer);
    procedure ClickFlecheDroite;
    procedure ClickFlecheGauche;
    procedure ClickFlecheTous;
    procedure ClickFlecheAucun;
    procedure ClickFlecheHaut;
    procedure ClickFlecheBas;
    procedure ClickFlecheHautMeth;
    procedure ClickFlecheBasMeth;
    procedure AjouteMethode;
    procedure SupprimeMethode;
    procedure ModifieMethode;
    procedure RefreshGrid(posListe, posAffiche: integer);
    procedure RefreshBouton;
    procedure RefreshGridMeth(posAffiche: integer);
    procedure RefreshBoutonMeth;
    procedure SetLastError(Num: integer; ou: string);
  end;

const
  BTN_DROIT = 'DROIT';
  BTN_GAUCHE = 'GAUCHE';
  BTN_HAUT = 'HAUT';
  BTN_BAS = 'BAS';
  BTN_TOUS = 'TOUS';
  BTN_AUCUN = 'AUCUN';
  BTN_NOUV = 'NOUV';
  BTN_SUPP = 'SUPP';
  BTN_MOD = 'MODIF';
  GRD_LISTE = 'GLISTE';
  GRD_AFFICHE = 'GAFFICHE';
  GRD_METHODE = 'GMETHODE';

  // libellés des messages
  TexteMessage: array[1..5] of string = (
    {1}'Vous devez indiquez une méthode de calcul de quantité à laisser au dépot',
    {2}'Vous devez renseigner une quantité à laisser au dépot',
    {3}'Vous devez renseigner un pourcentage de stock à laisser au dépot',
    {4}'Vous devez sélectionner au moins une Méthode',
    {5}'L''affectation de transfert générée est vide !!!'
    );

implementation

procedure TOF_MBOPROPTRF.OnArgument(St: string);
var critere: string;
begin
  inherited;
  Critere := (ReadTokenSt(St));
  if critere <> '' then SetControltext('CODPROPAFF', critere);
  GMETHODE := THGrid(GetControl('GMETHODE'));
  GAFFICHE := THGRID(GetControl('GAFFICHE'));
  GLISTE := THGRID(GetControl('GLISTE'));
  GAFFICHE.OnCellEnter := G_AFFCellEnter;
  GAFFICHE.OnCellExit := G_AFFCellExit;
  GAFFICHE.OnRowEnter := G_AFFRowEnter;
  GAFFICHE.OnRowExit := G_AFFRowExit;
  GLISTE.ColWidths[0] := 144;
  GLISTE.ColAligns[0] := taLeftJustify;
  GAFFICHE.ColWidths[0] := 244;
  GAFFICHE.ColAligns[0] := taLeftJustify;
  GAFFICHE.ColWidths[1] := 0;
  GAFFICHE.ColWidths[2] := 0;
  GAFFICHE.Options := GAFFICHE.Options + [goRowSelect];
  GAFFICHE.Options := GAFFICHE.Options - [goEditing];
  TobAffectationTrf := LaTOB; // Liste des articles sélectionnés dans le Multi-critères Article
  LaTOB := nil;
end;

procedure TOF_MBOPROPTRF.OnLoad;
var i: integer;
  Num, CodeMethode: string;
  Q, QM: TQuery;
  Etat, Reaffect: THRadioGroup;
begin
  inherited;
  TobMethode := TOB.Create('Methodes', nil, -1);
  TobListe := TOB.CREATE('Liste établissements', nil, -1);
  TobAffiche := TOB.CREATE('Etablissements affichés', nil, -1);
  Q := OpenSQL('select * from PROPTRANSFENT ' +
    'left join ETABLISS on ET_ETABLISSEMENT = GTE_DEPOTEMET ' +
    'where GTE_CODEPTRF = "' + GetControltext('CODPROPAFF') + '"', True);
  if not Q.EOF then
  begin
    SetControltext('PROPAFF', Q.Findfield('GTE_LIBPTRF').AsString);
    SetControltext('DEPOT', Q.Findfield('ET_LIBELLE').AsString);
    SetControltext('CODDEPOT', Q.Findfield('GTE_DEPOTEMET').AsString);
    Etat := THRadioGroup(GetControl('RGETAT'));
    if not (Q.Findfield('GTE_ENCOURSPTRF').AsString = 'X') then
      Etat.ItemIndex := 0
    else Etat.ItemIndex := 1;
    if (Q.Findfield('GTE_GENEREPTRF').AsString = 'X') then
      Etat.ItemIndex := 2;

    if (Q.Findfield('GTE_STOCKMIN').AsString) = '-' then
    begin
      SetControlChecked('STOCKMIN', False);
      SetControlText('STOCKMIN', '-');
    end else
    begin
      SetControlChecked('STOCKMIN', True);
      SetControlText('STOCKMIN', 'X');
    end;
    if (Q.Findfield('GTE_VIDDEPOT').AsString) = 'X' then
    begin
      SetControlChecked('VIDDEPOT', True);
      SetControlText('VIDDEPOT', 'X');
    end else
    begin
      SetControlChecked('VIDDEPOT', False);
      SetControlText('VIDDEPOT', '-');
    end;
    if (Q.Findfield('GTE_QTEFIXE').AsString) = 'X' then
    begin
      SetControlChecked('QTEFIXE', True);
      SetControlText('QTEFIXE', 'X');
      SetControlProperty('NBQTEFIXE', 'Value', Q.Findfield('GTE_NBQTEFIXE').AsInteger);
    end else
    begin
      SetControlChecked('QTEFIXE', False);
      SetControlText('QTEFIXE', '-');
      SetControlProperty('NBQTEFIXE', 'Value', 0);
    end;
    if (Q.Findfield('GTE_TXSTOCK').AsString) = 'X' then
    begin
      SetControlChecked('TXSTOCK', True);
      SetControlText('TXSTOCK', 'X');
      SetControlProperty('NBTXSTOCK', 'Value', Q.Findfield('GTE_NBTXSTOCK').AsInteger);
    end else
    begin
      SetControlChecked('TXSTOCK', False);
      SetControlText('TXSTOCK', '-');
      SetControlProperty('NBTXSTOCK', 'Value', 0);
    end;
    if (Q.Findfield('GTE_CONSULT').AsString) = 'X' then
    begin
      SetControlChecked('CONSULT', True);
      SetControlText('CONSULT', 'X');
    end else
    begin
      SetControlChecked('CONSULT', False);
      SetControlText('CONSULT', '-');
    end;
    if (Q.Findfield('GTE_ARRETRUPTURE').AsString) = '-' then
    begin
      SetControlChecked('ARRETRUPTURE', False);
      SetControlText('ARRETRUPTURE', '-');
    end else
    begin
      SetControlChecked('ARRETRUPTURE', True);
      SetControlText('ARRETRUPTURE', 'X');
    end;

    Reaffect := THRadioGroup(GetControl('REAFFECT'));
    if not (Q.Findfield('GTE_REAFFECTART').AsString = 'X') then
      Reaffect.ItemIndex := 0
    else Reaffect.ItemIndex := 1;
    if (Q.Findfield('GTE_REAFFECTBTQ').AsString = 'X') then
      Reaffect.ItemIndex := 2;

    if Q.Findfield('GTE_LIMITEART').AsString = 'AUC' then
      SetControlProperty('LIMITEART', 'ItemIndex', 0)
    else
      if Q.Findfield('GTE_LIMITEART').AsString = 'STK' then
      SetControlProperty('LIMITEART', 'ItemIndex', 1)
    else
      if Q.Findfield('GTE_LIMITEART').AsString = 'VTE' then
      SetControlProperty('LIMITEART', 'ItemIndex', 3)
    else SetControlProperty('LIMITEART', 'ItemIndex', 2);

    SetControlChecked('GTE_STKNET', Q.Findfield('GTE_STKNET').AsString = 'X');
    SetControlChecked('GTE_STKTRANSIT', Q.Findfield('GTE_STKTRANSIT').AsString = 'X');
    SetControlChecked('GTE_LIMITEENREG', Q.Findfield('GTE_LIMITEENREG').AsString = 'X');
    SetControlChecked('GTE_STKNETBTQ', Q.Findfield('GTE_STKNETBTQ').AsString = 'X');
    SetControlChecked('GTE_STKTRANSITBTQ', Q.Findfield('GTE_STKTRANSITBTQ').AsString = 'X');

    // Chargement des méthodes de la proposition
    for i := 1 to 10 do
    begin
      if (i < 10) then Num := IntToStr(i) else Num := 'A';
      CodeMethode := Q.Findfield('GTE_METHODE' + Num).AsString;
      if CodeMethode = '' then break;
      QM := OpenSQL('select * from PROPMETHODE ' +
        'where GTM_CODEMETPAFF="' + CodeMethode + '"', True);
      if not QM.EOF then
      begin
        AlimenteTobMethode(CodeMethode, QM.Findfield('GTM_TYPEMETPAFF').AsString,
          QM.Findfield('GTM_ARRONDI').AsString,
          QM.Findfield('GTM_UTILCOEFETAB').AsString,
          QM.Findfield('GTM_QTEDIM').AsFloat,
          QM.Findfield('GTM_QTEMINDIM').AsFloat,
          QM.Findfield('GTM_QTEMAXDIM').AsFloat,
          DatetoStr(QM.Findfield('GTM_DEBPERVTE').AsDateTime),
          DatetoStr(QM.Findfield('GTM_FINPERVTE').AsDateTime), TobMethode);
      end;
      Ferme(QM);
    end;
    if i > 1 then
    begin
      RefreshGridMeth(1);
      ChangeSelectBtq(0);
    end;
  end;
  Ferme(Q);
end;

procedure TOF_MBOPROPTRF.OnUpdate;
var ctrlzone, i: integer;
  ConsultPropo: boolean;
  Reaffect: THRadioGroup;
  TOBPropEnt, TobM: TOB;
  TOBProp: TOB;
  CodeDepotEmet, CodePropTrf, Num: string;
  Q: TQuery;
begin
  inherited;
  ctrlzone := 0;
  if TCheckbox(GetControl('STOCKMIN')).checked = True then
    inc(ctrlzone);
  if TCheckbox(GetControl('VIDDEPOT')).checked = True then
    inc(ctrlzone);
  if TCheckbox(GetControl('QTEFIXE')).checked = True then
    inc(ctrlzone);
  if TCheckbox(GetControl('TXSTOCK')).checked = True then
    inc(ctrlzone);

  if (ctrlzone <> 1) then
  begin
    SetLastError(1, '');
    exit;
  end;

  if (TCheckbox(GetControl('QTEFIXE')).checked = True)
    and (TSpinEdit(GetControl('NBQTEFIXE')).Value = 0) then
  begin
    SetLastError(2, 'NBQTEFIXE');
    exit;
  end;

  if (TCheckbox(GetControl('TXSTOCK')).checked = True)
    and (TSpinEdit(GetControl('NBTXSTOCK')).Value = 0) then
  begin
    SetLastError(3, 'NBTXSTOCK');
    exit;
  end;

  // Mémorisation du paramétrage de la proposition
  Q := OpenSQL('select * from PROPTRANSFENT ' +
    'where GTE_CODEPTRF = "' + GetControltext('CODPROPAFF') + '"', True);

  TOBPropEnt := TOB.Create('', nil, -1);
  TOBPropEnt.LoadDetailDB('PROPTRANSFENT', '', '', Q, False, True);
  Ferme(Q);
  if TOBPropEnt.Detail.Count > 0 then
  begin
    TOBPropEnt.Detail[0].PutValue('GTE_STOCKMIN', GetControlText('STOCKMIN'));
    TOBPropEnt.Detail[0].PutValue('GTE_VIDDEPOT', GetControlText('VIDDEPOT'));
    TOBPropEnt.Detail[0].PutValue('GTE_QTEFIXE', GetControlText('QTEFIXE'));
    TOBPropEnt.Detail[0].PutValue('GTE_TXSTOCK', GetControlText('TXSTOCK'));
    if GetControlText('QTEFIXE') = 'X' then
      TOBPropEnt.Detail[0].PutValue('GTE_NBQTEFIXE', TSpinEdit(GetControl('NBQTEFIXE')).Value)
    else TOBPropEnt.Detail[0].PutValue('GTE_NBQTEFIXE', 0);
    if GetControlText('TXSTOCK') = 'X' then
      TOBPropEnt.Detail[0].PutValue('GTE_NBTXSTOCK', TSpinEdit(GetControl('NBTXSTOCK')).Value)
    else TOBPropEnt.Detail[0].PutValue('GTE_NBTXSTOCK', 0);
    TOBPropEnt.Detail[0].PutValue('GTE_CONSULT', GetControlText('CONSULT'));
    TOBPropEnt.Detail[0].PutValue('GTE_ARRETRUPTURE', GetControlText('ARRETRUPTURE'));
    Reaffect := THRadioGroup(GetControl('REAFFECT'));
    if Reaffect.ItemIndex <> 0 then
      TOBPropEnt.Detail[0].PutValue('GTE_REAFFECTART', 'X')
    else TOBPropEnt.Detail[0].PutValue('GTE_REAFFECTART', '-');
    if Reaffect.ItemIndex = 2 then
      TOBPropEnt.Detail[0].PutValue('GTE_REAFFECTBTQ', 'X')
    else TOBPropEnt.Detail[0].PutValue('GTE_REAFFECTBTQ', '-');

    TOBPropEnt.Detail[0].PutValue('GTE_LIMITEART', THRadioGroup(GetControl('LIMITEART')).Value);
    TOBPropEnt.Detail[0].PutValue('GTE_STKNET', GetControlText('GTE_STKNET'));
    TOBPropEnt.Detail[0].PutValue('GTE_STKTRANSIT', GetControlText('GTE_STKTRANSIT'));
    TOBPropEnt.Detail[0].PutValue('GTE_STKNETBTQ', GetControlText('GTE_STKNETBTQ'));
    TOBPropEnt.Detail[0].PutValue('GTE_STKTRANSITBTQ', GetControlText('GTE_STKTRANSITBTQ'));
    TOBPropEnt.Detail[0].PutValue('GTE_LIMITEENREG', GetControlText('GTE_LIMITEENREG'));

    // Récupération des différentes méthodes sélectionnées
    for i := 0 to TobMethode.Detail.Count - 1 do
    begin
      TobM := TobMethode.Detail[i];
      if i = 10 then break;
      if (i < 9) then Num := IntToStr(i + 1) else Num := 'A';
      TOBPropEnt.Detail[0].PutValue('GTE_METHODE' + Num, TobM.GetValue('CODE'));
    end;
    for i := TobMethode.Detail.Count to 9 do
    begin
      if (i < 9) then Num := IntToStr(i + 1) else Num := 'A';
      TOBPropEnt.Detail[0].PutValue('GTE_METHODE' + Num, '');
    end;

    TOBPropEnt.Detail[0].PutValue('GTE_DATEMODIF', Date);
    TOBPropEnt.Detail[0].PutValue('GTE_UTILISATEUR', V_PGI.User);
    TOBPropEnt.InsertOrUpdateDB(False);
  end;
  TOBPropEnt.Free;

  // ----------------------------------------------------------------------------------------
  // Lancement du traitement de calcul de la nouvelle affectation de proposition de transfert.
  // ----------------------------------------------------------------------------------------
  if HShowMessage('0;Confirmation;Confirmez-vous la génération de cette nouvelle affectation?;Q;YN;N;N;', '', '')
    <> mrYes then
    begin
      TForm(Ecran).ModalResult := 0;
      exit;
    end;
  if GetCheckBoxState('PLUSTARD') = cbChecked then
  //if TCheckBox(FF.FindComponent('PLUSTARD')).Checked then
  begin
    if not FOLanceDiffere('Génération d''affectation') then
    begin
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  end;
  TOBProp := TraitementAffectationTrf(TForm(Ecran), TobAffectationTrf, TobMethode, TobAffiche, 'A');

  if (TOBProp = nil) or (TOBProp.detail.count = 0) then
  begin
    if TOBProp <> nil then PgiInfo(TexteMessage[5], Ecran.Caption);
  end else
  begin
    ConsultPropo := TCheckbox(GetControl('CONSULT')).checked;
    if (ConsultPropo) then
    begin
      // Consultation et/ou modification de la proposition d'affectation du transfert
      TheTOB := TOBProp;
      CodeDepotEmet := GetControlText('CODDEPOT');
      AGLLanceFiche('MBO', 'PR_AFFTRANSFERT', '', '', 'ACTION=MODIFICATION;CODEDEPOTEMET=' + CodeDepotEmet);
      TOBProp := TheTOB;
      TheTOB := nil;
    end;
  end;

  // Enregistrement de la proposition d'affectation du transfert dans la Table
  if (TOBProp <> nil) and (TOBProp.Detail.Count > 0) then
  begin
    if THRadioGroup(GetControl('RGETAT')).ItemIndex <> 2 then
    begin
      TOBProp.SetAllModifie(True);
      if THRadioGroup(GetControl('RGETAT')).ItemIndex = 0 then
        TOBProp.InsertDB(nil, True)
      else TOBProp.InsertOrUpdateDB(True);
      CodePropTrf := GetControlText('CODPROPAFF');
      ExecuteSQL('Update PROPTRANSFENT Set GTE_ENCOURSPTRF="X" where GTE_CODEPTRF="' + CodePropTrf + '"');
    end;
  end;

  TOBProp.Free;
end;

procedure TOF_MBOPROPTRF.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
  TForm(Ecran).ModalResult := 0;
end;

procedure TOF_MBOPROPTRF.OnClose;
begin
  inherited;
  TobMethode.Free;
  TobMethode := nil;
  TheTob := nil;
  TobAffiche.free;
  TobAffiche := nil;
  TobListe.free;
  TobListe := nil;
  TobAffectationTrf.Free;
  TobAffectationTrf := nil;
end;

procedure TOF_MBOPROPTRF.RefreshGrid(posListe, posAffiche: integer);
begin
  TobAffiche.PutGridDetail(GAFFICHE, False, False, 'LIBELLE;POIDS;COEFF', True);
  TobListe.PutGridDetail(GLISTE, False, False, 'LIBELLE', True);
  GAFFICHE.Row := Min(posAffiche, GAFFICHE.RowCount - 1);
  GLISTE.Row := Min(posListe, GLISTE.RowCount - 1);
  RefreshBouton;
end;

// Boutons enable / disable

procedure TOF_MBOPROPTRF.RefreshBouton;
begin
  SetControlEnabled('BFLECHEDROITE', TobListe.Detail.Count > 0);
  SetControlEnabled('BFLECHEGAUCHE', TobAffiche.Detail.Count > 0);
  SetControlEnabled('BFLECHEHAUT', GAFFICHE.Row > 1);
  SetControlEnabled('BFLECHEBAS', GAFFICHE.Row < GAFFICHE.RowCount - 1);
  SetControlEnabled('BFLECHETOUS', TobListe.Detail.Count > 0);
  SetControlEnabled('BFLECHEAUCUN', TobAffiche.Detail.Count > 0);
end;

procedure TOF_MBOPROPTRF.RefreshGridMeth(posAffiche: integer);
begin
  TobMethode.PutGridDetail(GMETHODE, False, False, 'ORDRE;LIBELLE', True);
  GMETHODE.Row := Min(posAffiche, GMETHODE.RowCount - 1);
  RefreshBoutonMeth;
end;

// Boutons enable / disable

procedure TOF_MBOPROPTRF.RefreshBoutonMeth;
begin
  SetControlEnabled('BFHAUT', GMETHODE.Row > 1);
  SetControlEnabled('BFBAS', GMETHODE.Row < GMETHODE.RowCount - 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheDroite;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if GLISTE.Row < 0 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if TobAffiche.Detail.Count > 0 then indiceFille := GAFFICHE.Row else indiceFille := 0;
  TobListe.detail[GLISTE.Row - 1].ChangeParent(TobAffiche, indiceFille);
  RefreshGrid(GLISTE.Row, GAFFICHE.Row + 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheGauche;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if GAFFICHE.Row < 0 then exit;
  // Changement du parent de l'élément des établissements affichés
  if TobListe.Detail.Count > 0 then indiceFille := GLISTE.Row else indiceFille := 0;
  TobAffiche.detail[GAFFICHE.Row - 1].ChangeParent(TobListe, indiceFille);
  RefreshGrid(GLISTE.Row + 1, GAFFICHE.Row);
end;

procedure TOF_MBOPROPTRF.ClickFlecheTous;
var indiceFille, iGrd, posliste: integer;
begin
  if GLISTE.RowCount < 2 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if GAFFICHE.RowCount > 2 then indiceFille := GAFFICHE.Row else indiceFille := 0;
  posliste := TobListe.detail.count - 1;
  for iGrd := 0 to posliste do TobListe.detail[0].ChangeParent(TobAffiche, indiceFille + iGrd);
  RefreshGrid(1, indiceFille + posliste);
end;

procedure TOF_MBOPROPTRF.ClickFlecheAucun;
var indiceFille, iGrd, posaffiche: integer;
begin
  if GAFFICHE.RowCount < 2 then exit;
  // Changement du parent de l'élément de la liste des établissements
  posaffiche := TobAffiche.detail.count - 1;
  if GLISTE.RowCount > 2 then indiceFille := GLISTE.Row else indiceFille := 0;
  for iGrd := 0 to posaffiche do TobAffiche.detail[0].ChangeParent(TobListe, indiceFille + iGrd);
  RefreshGrid(indiceFille + posaffiche, 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheHaut;
begin
  if GAFFICHE.Row < 1 then exit;
  // Changement de l'indice dans la Tob parent
  TobAffiche.detail[GAFFICHE.Row - 1].ChangeParent(TobAffiche, GAFFICHE.Row - 2);
  RefreshGrid(GLISTE.Row, GAFFICHE.Row - 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheBas;
begin
  if GAFFICHE.Row > GAFFICHE.RowCount - 2 then exit;
  // Changement de l'indice dans la Tob parent
  TobAffiche.detail[GAFFICHE.Row - 1].ChangeParent(TobAffiche, GAFFICHE.Row);
  RefreshGrid(GLISTE.Row, GAFFICHE.Row + 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheHautMeth;
begin
  if GMETHODE.Row < 1 then exit;
  // Changement de l'indice dans la Tob parent
  TobMethode.detail[GMETHODE.Row - 1].PutValue('ORDRE', GMETHODE.Row - 1);
  TobMethode.detail[GMETHODE.Row - 2].PutValue('ORDRE', GMETHODE.Row);
  TobMethode.detail[GMETHODE.Row - 1].ChangeParent(TobMethode, GMETHODE.Row - 2);
  RefreshGridMeth(GMETHODE.Row - 1);
end;

procedure TOF_MBOPROPTRF.ClickFlecheBasMeth;
begin
  if GMETHODE.Row > GMETHODE.RowCount - 2 then exit;
  // Changement de l'indice dans la Tob parent
  TobMethode.detail[GMETHODE.Row - 1].PutValue('ORDRE', GMETHODE.Row + 1);
  TobMethode.detail[GMETHODE.Row].PutValue('ORDRE', GMETHODE.Row);
  TobMethode.detail[GMETHODE.Row - 1].ChangeParent(TobMethode, GMETHODE.Row);
  RefreshGridMeth(GMETHODE.Row + 1);
end;

procedure TOF_MBOPROPTRF.G_AFFCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if GAFFICHE.Col = 0 then GAFFICHE.Col := 1;
  if (TobMethode.Detail.count > 0) then
  begin
    if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') = 'PDS') and (GAFFICHE.Col = 2) then GAFFICHE.Col := 1;
    if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') <> 'PDS') and (GAFFICHE.Col = 1) then GAFFICHE.Col := 2;
  end;
end;

procedure TOF_MBOPROPTRF.G_AFFCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (Acol = 1) then TobAffiche.detail[ARow - 1].PutValue('POIDS', ValeurI(GAFFICHE.Cells[Acol, ARow]));
  if (Acol = 2) then TobAffiche.detail[ARow - 1].PutValue('COEFF', Valeur(GAFFICHE.Cells[Acol, ARow]));
  if (TobMethode.Detail.count > 0) then
    if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') = 'PDS') then
      Calculcoeff;
end;

procedure TOF_MBOPROPTRF.Calculcoeff;
var TOBT: TOB;
  i_ind, totpoids: integer;
  totcoeff: double;
begin
  totpoids := 0;
  for i_ind := 0 to TobAffiche.Detail.count - 1 do
  begin
    TobT := TobAffiche.Detail[i_ind];
    totpoids := totpoids + TOBT.GetValue('POIDS');
  end;
  for i_ind := 0 to TobAffiche.Detail.count - 1 do
  begin
    TobT := TobAffiche.Detail[i_ind];
    if TOBT.GetValue('POIDS') <> 0 then
      TOBT.PutValue('COEFF', Arrondi((TOBT.GetValue('POIDS') / totpoids) * 100.0, 2))
    else
      TOBT.PutValue('COEFF', 0.0);
  end;
  totcoeff := 0;
  for i_ind := 0 to TobAffiche.Detail.count - 1 do
  begin
    TobT := TobAffiche.Detail[i_ind];
    totcoeff := totcoeff + TOBT.GetValue('COEFF');
  end;
  if totcoeff <> 100.0 then
  begin
    totcoeff := 100.0 - totcoeff;
    for i_ind := 0 to TobAffiche.Detail.count - 1 do
    begin
      if TobAffiche.Detail[i_ind].GetValue('POIDS') <> 0 then
      begin
        TobAffiche.Detail[i_ind].PutValue('COEFF', (TobAffiche.Detail[i_ind].GetValue('COEFF') + totcoeff));
        break;
      end;
    end;
  end;
  RefreshGrid(GLISTE.Row, GAFFICHE.Row);
end;

procedure TOF_MBOPROPTRF.G_AFFRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_MBOPROPTRF.G_AFFRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_MBOPROPTRF.AjouteMethode;
begin
  TheTob := TobMethode;
  AglLanceFiche('MBO', 'PR_METHAFFECT', '', '', GetControltext('CODDEPOT'));
  TobMethode := TheTob;
  TobMethode.PutGridDetail(GMETHODE, False, False, 'ORDRE;LIBELLE', True);
  if TobMethode.detail.count > 0 then
  begin
    GMETHODE.Row := TobMethode.detail.count;
    ChangeSelectBtq(GMETHODE.Row - 1);
  end;
  RefreshBoutonMeth;
end;

procedure TOF_MBOPROPTRF.SupprimeMethode;
var i_ind: integer;
begin
  if GMETHODE.Row > GMETHODE.RowCount - 1 then exit;
  // Changement de l'indice dans la Tob parent
  TobMethode.detail[GMETHODE.Row - 1].Free;
  for i_ind := 0 to TobMethode.detail.count - 1 do
  begin
    TobMethode.detail[i_ind].PutValue('ORDRE', i_ind + 1);
  end;
  RefreshGridMeth(GMETHODE.Row);
  ChangeSelectBtq(GMETHODE.Row - 1);
end;

procedure TOF_MBOPROPTRF.ModifieMethode;
begin
  if GMETHODE.Row > GMETHODE.RowCount - 1 then exit;
  ChangeSelectBtq(GMETHODE.Row - 1);
end;

procedure TOF_MBOPROPTRF.ChangeSelectBtq(Tobind: integer);
var iListe, iEtab, i_ind: integer;
  bTrouve: boolean;
  QQ, QQ2: TQuery;
  TobEtab, TobTemp, TobT: TOB;
begin
  inherited;
  if not (ctxMode in V_PGI.PGIContexte) then exit;
  // Chargement de la liste des établissements et du paramétrage existant.
  TobListe.ClearDetail;
  TobAffiche.ClearDetail;
  TobEtab := TOB.CREATE('LEtab', nil, -1);
  TobTemp := TOB.Create('', nil, -1);
  // Chargement paramétrage sauvegardé
  if TobMethode.Detail.count > 0 then
  begin
    QQ2 := OpenSQL('select GTQ_ETABLISSEMENT,ET_LIBELLE,GTQ_COEFREPAR,GTQ_POIDS from PROPMETETAB ' +
      'left join ETABLISS on ET_ETABLISSEMENT = GTQ_ETABLISSEMENT ' +
      'where GTQ_CODEMETPAFF = "' + TobMethode.Detail[Tobind].GetValue('CODE') + '" order by GTQ_LIGMETPAFF', True);
    if not QQ2.eof then TobTemp.LoadDetailDB('', '', '', QQ2, false);

    for i_ind := 0 to TobTemp.Detail.count - 1 do
    begin
      TobT := TOB.Create('SelEtablissement', TobEtab, -1);
      TobT.AddChampSup('ETABLISSEMENT', False);
      TobT.AddChampSup('LIBELLE', False);
      TobT.AddChampSup('COEFF', False);
      TobT.AddChampSup('POIDS', False);
      TobT.InitValeurs;
      TobT.PutValue('ETABLISSEMENT', TobTemp.Detail[i_ind].GetValue('GTQ_ETABLISSEMENT'));
      TobT.PutValue('LIBELLE', TobTemp.Detail[i_ind].GetValue('ET_LIBELLE'));
      TobT.PutValue('COEFF', TobTemp.Detail[i_ind].GetValue('GTQ_COEFREPAR'));
      TobT.PutValue('POIDS', TobTemp.Detail[i_ind].GetValue('GTQ_POIDS'));
    end;
    TobTemp.free;
    Ferme(QQ2);

    TobTemp := TOB.Create('', nil, -1);
    QQ := OpenSQL('select ET_ETABLISSEMENT,ET_LIBELLE,ET_ABREGE from ETABLISS order by ET_ETABLISSEMENT', True);
    if not QQ.eof then TobTemp.LoadDetailDB('', '', '', QQ, false);
    for i_ind := 0 to TobTemp.Detail.count - 1 do
    begin
      //if (TobTemp.Detail[i_ind].GetValue('ET_ETABLISSEMENT') <> GetControltext('CODDEPOT')) then
      //  begin
      TobT := TOB.Create('Etablissement', TobListe, -1);
      TobT.AddChampSup('ETABLISSEMENT', False);
      TobT.AddChampSup('LIBELLE', False);
      TobT.AddChampSup('COEFF', False);
      TobT.AddChampSup('POIDS', False);
      TobT.InitValeurs;
      TobT.PutValue('ETABLISSEMENT', TobTemp.Detail[i_ind].GetValue('ET_ETABLISSEMENT'));
      TobT.PutValue('LIBELLE', TobTemp.Detail[i_ind].GetValue('ET_LIBELLE'));
      TobT.PutValue('COEFF', 1.0);
      TobT.PutValue('POIDS', 0);
      //  end;
    end;
    TobTemp.free;
    Ferme(QQ);

    if TobEtab.Detail.Count > 0 then
    begin
      // Chargement de la liste des établissements dans TobListe
      // et bascule des établissements de TobEtat dans TobAffiche trié suivant l'ordre définit dans TobEtat
      iEtab := 0;
      while iEtab < TobEtab.Detail.Count do
      begin
        iListe := 0;
        bTrouve := False;
        while (not bTrouve) and (iListe < TobListe.Detail.Count) do
        begin
          bTrouve := (TobListe.Detail[iListe].GetValue('ETABLISSEMENT')
            = TobEtab.Detail[iEtab].GetValue('ETABLISSEMENT'));
          inc(iListe);
        end;
        if bTrouve then // Transfert dans TobAffiche
        begin
          TobListe.Detail[iListe - 1].PutValue('COEFF', TobEtab.Detail[iEtab].GetValue('COEFF'));
          TobListe.Detail[iListe - 1].PutValue('POIDS', TobEtab.Detail[iEtab].GetValue('POIDS'));
          TobListe.Detail[iListe - 1].ChangeParent(TobAffiche, -1);
        end;
        inc(iEtab);
      end;
    end else
    begin
      // Paramétrage non prédéfini, chargement dans TobAffiche trié par code établissement
      for iListe := 0 to TobListe.Detail.Count - 1 do
      begin
        TobListe.detail[0].ChangeParent(TobAffiche, iListe);
      end;
    end;
  end;
  // Affichage des tobs
  GLISTE := THGrid(GetControl('GLISTE'));
  GAFFICHE := THGrid(GetControl('GAFFICHE'));
  TobAffiche.PutGridDetail(GAFFICHE, False, False, 'LIBELLE;POIDS;COEFF', True);
  TobListe.PutGridDetail(GLISTE, False, False, 'LIBELLE', True);
  setcontroltext('ARRONDI', TobMethode.Detail[Tobind].GetValue('ARRONDI'));
  Changetaillecol;
  RefreshBouton;
  TobEtab.free;
end;

procedure TOF_MBOPROPTRF.Changetaillecol;
begin
  GLISTE.ColWidths[0] := 144;
  GLISTE.ColAligns[0] := taLeftJustify;
  GAFFICHE.ColWidths[0] := 144;
  GAFFICHE.ColAligns[0] := taLeftJustify;
  GAFFICHE.ColWidths[1] := 50;
  GAFFICHE.ColAligns[1] := taRightJustify;
  GAFFICHE.ColWidths[2] := 50;
  GAFFICHE.ColAligns[2] := taRightJustify;
  GAFFICHE.Col := 1;
  if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') <> 'PDS') then
  begin
    GAFFICHE.Col := 2;
    GAFFICHE.ColWidths[0] := 194;
    GAFFICHE.ColWidths[1] := 0;
  end;
  if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('UTILCOEF') = '-')
    and (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') <> 'PDS') then
  begin
    GAFFICHE.ColWidths[0] := 244;
    GAFFICHE.ColWidths[1] := 0;
    GAFFICHE.ColWidths[2] := 0;
  end;
  if (GAFFICHE.ColWidths[1] = 0) and (GAFFICHE.ColWidths[2] = 0) then
  begin
    GAFFICHE.Options := GAFFICHE.Options + [goRowSelect];
    GAFFICHE.Options := GAFFICHE.Options - [goEditing];
  end else
  begin
    GAFFICHE.Options := GAFFICHE.Options - [goRowSelect];
    GAFFICHE.Options := GAFFICHE.Options + [goEditing];
  end;
  GAFFICHE.ColFormats[1] := '##';
  GAFFICHE.ColFormats[2] := '# ##0.00';
end;

procedure AGLOnMajListeBtq(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTof
  else exit;
  if (TOTOF is TOF_MBOPROPTRF) then TOF_MBOPROPTRF(TOTOF).OnMajListeBtq else exit;
end;

procedure TOF_MBOPROPTRF.OnMajListeBtq;
var stReq: string;
  iAffiche, indiceFille: integer;
  TOBT: TOB;
  maj: boolean;
begin
  if ctxMode in V_PGI.PGIContexte then
  begin
    stReq := 'delete from PROPMETETAB where GTQ_CODEMETPAFF="' + TobMethode.Detail[GMETHODE.Row - 1].GetValue('CODE') + '"';
    ExecuteSQL(stReq);
    // Svgde des établissements affichés en cconsultation multi-dépôts : position > 100
    iAffiche := 0;
    while iAffiche < TobAffiche.Detail.count do
    begin
      TobT := TobAffiche.Detail[iAffiche];
      maj := True;
      if (TobMethode.Detail[GMETHODE.Row - 1].GetValue('TYPE') = 'PDS')
        and (TobAffiche.Detail[iAffiche].GetValue('POIDS') = 0) then maj := False;

      if maj then
      begin
        stReq := 'insert into PROPMETETAB (GTQ_CODEMETPAFF,GTQ_LIGMETPAFF,GTQ_ETABLISSEMENT,GTQ_COEFREPAR,GTQ_POIDS)' +
          ' values ("' + TobMethode.Detail[GMETHODE.Row - 1].GetValue('CODE') + '",' + inttostr(iAffiche) +
          ',"' + TobT.GetValue('ETABLISSEMENT') + '",' + StrfPoint(TobT.GetValue('COEFF')) + ',' + inttostr(TobT.GetValue('POIDS')) + ')';
        ExecuteSQL(stReq);
      end else
      begin
        if TobListe.Detail.Count > 0 then indiceFille := GLISTE.Row else indiceFille := 0;
        TobT.ChangeParent(TobListe, indiceFille);
        RefreshGrid(GLISTE.Row + 1, GAFFICHE.Row);
        iAffiche := iAffiche - 1;
      end;
      inc(iAffiche);
    end;

    // Svgde de la méthode d'arrondi
    stReq := 'Update PROPMETHODE Set GTM_ARRONDI="' + GetControlText('ARRONDI') + '" ' +
      'where GTM_CODEMETPAFF="' + TobMethode.Detail[GMETHODE.Row - 1].GetValue('CODE') + '"';
    ExecuteSQL(stReq);
    TobMethode.Detail[GMETHODE.Row - 1].PutValue('ARRONDI', GetControlText('ARRONDI'));
  end;
end;

procedure AGLOnClickBouton3(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOPROPTRF) then
  begin
    if Parms[1] = BTN_DROIT then TOF_MBOPROPTRF(TOTOF).ClickFlecheDroite
    else if Parms[1] = BTN_GAUCHE then TOF_MBOPROPTRF(TOTOF).ClickFlecheGauche
    else if Parms[1] = BTN_HAUT then TOF_MBOPROPTRF(TOTOF).ClickFlecheHaut
    else if Parms[1] = BTN_BAS then TOF_MBOPROPTRF(TOTOF).ClickFlecheBas
    else if Parms[1] = BTN_TOUS then TOF_MBOPROPTRF(TOTOF).ClickFlecheTous
    else if Parms[1] = BTN_AUCUN then TOF_MBOPROPTRF(TOTOF).ClickFlecheAucun
    else if (Parms[1] = GRD_LISTE) or (Parms[1] = GRD_AFFICHE) then TOF_MBOPROPTRF(TOTOF).RefreshBouton;
  end;
end;

procedure AGLOnClickBoutMeth(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_MBOPROPTRF) then
  begin
    if Parms[1] = BTN_NOUV then TOF_MBOPROPTRF(TOTOF).AjouteMethode
    else if Parms[1] = BTN_SUPP then TOF_MBOPROPTRF(TOTOF).SupprimeMethode
    else if Parms[1] = BTN_MOD then TOF_MBOPROPTRF(TOTOF).ModifieMethode
    else if Parms[1] = BTN_HAUT then TOF_MBOPROPTRF(TOTOF).ClickFlecheHautMeth
    else if Parms[1] = BTN_BAS then TOF_MBOPROPTRF(TOTOF).ClickFlecheBasMeth
    else if (Parms[1] = GRD_METHODE) then
    begin
      TOF_MBOPROPTRF(TOTOF).ModifieMethode;
      TOF_MBOPROPTRF(TOTOF).RefreshBoutonMeth;
    end;
  end;
end;

initialization
  registerclasses([TOF_MBOPROPTRF]);
  RegisterAglProc('OnClickBouton3', True, 1, AGLOnClickBouton3);
  RegisterAglProc('OnClickBoutmeth', True, 1, AGLOnClickBoutmeth);
  RegisterAglProc('OnMajListeBtq', True, 1, AGLOnMajListeBtq);
end.
