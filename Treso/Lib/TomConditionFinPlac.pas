{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               23/10/01  BT   Création de l'unité
               03/08/03  JP   Migration en eAGL
 1.20.001.001  22/03/04  JP   Petites corrections dans MajEnSerie (LastError)
 6.20.001.002  19/01/05  JP   Gestion d'une duplication digne de ce nom !!
 6.51.001.001  15/11/05  JP   FQ 10307 : On court-circuite le onChange si en mode consultation
 7.05.001.001  23/10/06  JP   Gestion des filtres multi sociétés
--------------------------------------------------------------------------------------}
unit TomConditionFinPlac ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTob,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Fiche, HDB,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, SysUtils, HCtrls, HEnt1, UTOM,
  Spin;

type
  TOM_CONDITIONFINPLAC = class (TOM)
    procedure OnNewRecord               ; override;
    procedure OnUpdateRecord            ; override;
    procedure OnArgument    (S : string); override;
    procedure OnLoadRecord              ; override;
  protected
    {$IFDEF EAGLCLIENT}
    edCompte : THEdit;
    cbBanque : THValComboBox;
    {$ELSE}
    edCompte : THDBEdit;
    cbBanque : THDBValComboBox;
    {$ENDIF EAGLCLIENT}

    cbTauxFixe       : TCheckBox;
    edTauxVar        : THEdit;
    edValTaux        : THEdit;
    edValMultiple    : THEdit;
    edValMajoration  : THEdit;
    cbTauxPrecompte  : TCheckBox;
    spNbJourBanque   : THSpinEdit;
    spNbJourMinAgios : THSpinEdit;
    spNbJourEncaiss  : THSpinEdit;
    cbBaseCalcul     : THValComboBox;
    edDateDebut      : THEdit;
    edDateFin        : THEdit;
    TousCpt          : TCheckBox;

    Devise           : string;
    FValiderClick    : TNotifyEvent;
    FOnCloseQuery    : TCloseQueryEvent;
    General          : string;
    Condition        : string;
    DuplicationOk    : Boolean; {19/01/05}

    procedure CompteOnChange    (Sender : TObject);
    procedure DeviseOnChange    (Sender : TObject);
    procedure BanqueOnChange    (Sender : TObject);
    procedure TauxFixeOnClick   (Sender : TObject);
    procedure NbJourMiniOnChange(Sender : TObject);
    procedure TraiterParAgence  (Sender : TObject);
    procedure MajEnSerie        (Sender : TObject);
    procedure MajUnSeul         (Sender : TObject);
    procedure FicheCloseQuery   (Sender : TObject; var CanClose: Boolean);
    procedure GestionDevise     ;
  end ;

procedure TRLanceFiche_ConditionsFP(Dom, Fiche, Range, Lequel, Arguments : string);

Implementation

uses
  Commun, ExtCtrls, Constantes, HMsgBox, ParamSoc, UObjGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ConditionsFP(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  ReadTokenSt(S);
  Condition := ReadTokenSt(S);
  General   := ReadTokenSt(S);
  {19/01/05 : Correction de la duplication}
  DuplicationOk := ReadTokenSt(S) = CODEDUPLICAT;

  cbTauxFixe := TCheckBox(GetControl('TCF_TAUXFIXE'));
  edTauxVar := THEdit(GetControl('TCF_TAUXVAR'));
  edValTaux := THEdit(GetControl('TCF_VALTAUX'));
  edValMultiple := THEdit(GetControl('TCF_VALMULTIPLE'));
  edValMajoration := THEdit(GetControl('TCF_VALMAJORATION'));
  cbTauxPrecompte := TCheckBox(GetControl('TCF_TAUXPRECOMPTE'));
  cbBaseCalcul := THValComboBox(GetControl('TCF_BASECALCUL'));
  spNbJourBanque := THSpinEdit(GetControl('TCF_NBJOURBANQUE'));
  spNbJourMinAgios := THSpinEdit(GetControl('TCF_NBJOURMINAGIOS'));
  spNbJourEncaiss := THSpinEdit(GetControl('TCF_NBJOURENCAISS'));
  edDateDebut := THEdit(GetControl('TCF_DATEDEBUT'));
  edDateFin := THEdit(GetControl('TCF_DATEFIN'));
  {$IFDEF EAGLCLIENT}
  edCompte := THEdit(GetControl('TCF_GENERAL'));
  cbBanque := THValComboBox(GetControl('TCF_BANQUE'));
  {$ELSE}
  edCompte := THDBEdit(GetControl('TCF_GENERAL'));
  cbBanque := THDBValComboBox(GetControl('TCF_BANQUE'));
  {$ENDIF EAGLCLIENT}

  TousCpt := TCheckBox(GetControl('CKAGENCE'));
  TousCpt.OnClick := TraiterParAgence;
  {La case à cocher n'est active qu'en création}
  TGroupBox(GetControl('GBCOLLECT')).Enabled := TFFiche(Ecran).TypeAction in [taCreat, taCreatEnSerie, taCreatOne];
  TousCpt.Enabled := TGroupBox(GetControl('GBCOLLECT')).Enabled;
  {On ne peut modifier la banque qu'en création, puisqu'en consultation on ne peut modifier le
   compte qui fait partie de l'index}
  SetControlEnabled('TCF_BANQUE', TousCpt.Enabled);
  cbBanque.OnChange := BanqueOnChange;
  SetPlusBancaire(cbBanque, 'PQ', CODECOURANTS + ';' + CODETITRES + ';');

  GestionDevise;

  edCompte.OnChange := CompteOnChange;
  cbTauxFixe.OnClick := TauxFixeOnClick;
  spNbJourBanque.OnChange := NbJourMiniOnChange;
  spNbJourMinAgios.OnChange := NbJourMiniOnChange;
  spNbJourEncaiss.OnChange := NbJourMiniOnChange;

  FValiderClick := TFFiche(Ecran).BValider.OnClick;
  FOnCloseQuery := Ecran.OnCloseQuery;
  TFFiche(Ecran).BValider.OnClick := MajUnSeul;
  THValComboBox(GetControl('CBDEVISE')).OnChange := DeviseOnChange;

  SetControlProperty('PAGES', 'TABWIDTH', GetControl('PAGES').Width - 4);
  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.CompteOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  GestionDevise;
end;

 // RaZ et désactivation du taux fixe ou variable
{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.TauxFixeOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Bool: Boolean;
begin
  Bool := cbTauxFixe.Checked;
  if not Bool then begin
    {Si à taux variable, on réinitialise les zones concernant le taux fixe}
    if GetField('TCF_VALTAUX') <> 0 then SetField('TCF_VALTAUX', '0');
    if GetField('TCF_TAUXPRECOMPTE') = 'X' then SetField('TCF_TAUXPRECOMPTE', '-');
    edValTaux.Text := '0';
    cbTauxPrecompte.Checked := False;
  end
  else begin
    {Si à taux fixe, on réinitialise les zones concernant le taux variable}
    if GetField('TCF_TAUXVAR') <> '' then SetField('TCF_TAUXVAR', '');
    if GetField('TCF_VALMULTIPLE') <> '0' then SetField('TCF_VALMULTIPLE', '0');
    if GetField('TCF_VALMAJORATION') <> '0' then SetField('TCF_VALMAJORATION', '0');
    edTauxVar.Text := '';
    edValMultiple.Text := '0';
    edValMajoration.Text := '0';
  end;
  
  edValTaux.Enabled       := Bool;
  //cbTauxPrecompte.Enabled := Bool;
  edTauxVar.Enabled       := not Bool;
  edValMultiple.Enabled   := not Bool;
  edValMajoration.Enabled := not Bool;
  SetControlEnabled('FIXE', Bool);
  SetControlEnabled('TCF_TAUXPRECOMPTE', Bool);
  SetControlEnabled('VARIABLE', not Bool);
  SetControlEnabled('TTCF_VALMULTIPLE', not Bool);
  SetControlEnabled('TTCF_VALMAJORATION', not Bool);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.NbJourMiniOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if TSpinEdit(Sender).Value < 0 then
    TSpinEdit(Sender).Value := 0;
end;


{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {06/08/04 : FQ 10043 afficher le libellé plutôt que le compte général}
  Ecran.Caption := 'Conditions de financement et de placement du compte ' + RechDom('TRBANQUECP', GetField('TCF_GENERAL'), False);
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  // Vérification avant de valider
  if not cbTauxFixe.Checked and (edTauxVar.Text = '') then begin
    TrShowMessage(Ecran.Caption, 3, 'Taux variable', '');
    edTauxVar.SetFocus;
    LastError := -1;
    Exit;
  end
  else if cbTauxFixe.Checked and (edValTaux.Text = '') then begin
    TrShowMessage(Ecran.Caption, 3, 'Taux fixe', '');
    edTauxVar.SetFocus;
    LastError := -1;
    Exit;
  end;

  if cbBaseCalcul.Value = '' then
  begin
    TrShowMessage(Ecran.Caption, 3, 'Base de calcul', '');
    cbBaseCalcul.SetFocus;
    LastError := -1;
    Exit;
  end;

  if StrToDateTime(edDateDebut.Text) > StrToDateTime(edDateFin.Text) then
  begin
    TrShowMessage(Ecran.Caption, 4, '', '');
    edDateDebut.SetFocus;
    LastError := -1;
  end;

  if (GetField('TCF_GENERAL') = '') then begin
    LastError := 1;
    HShowMessage('2;' + Ecran.Caption + ';Veuillez renseigner le compte bancaire !;W;O;O;O;', '', '');
    edCompte.SetFocus;
    Exit;
  end;

  if GetControlText('TCF_CONDITIONFP') = '' then begin
    LastError := 1;
    HShowMessage('2;' + Ecran.Caption + ';Veuillez sélectionner une transaction !;W;O;O;O;', '', '');
    SetFocusControl('TCF_CONDITIONFP');
    Exit;
  end;

  if GetControlText('TCF_BANQUE') = '' then begin
    LastError := 1;
    HShowMessage('2;' + Ecran.Caption + ';Veuillez sélectionner une banque !;W;O;O;O;', '', '');
    SetFocusControl('TCF_BANQUE');
    Exit;
  end;
end;

{Surcharge du FormCloseQuery lors d'une création en série pour éviter le message "Voulez-
 cous enregistrer vos modification
{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.FicheCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := LastError = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.TraiterParAgence(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not TGroupBox(GetControl('GBCOLLECT')).Enabled then Exit;
  SetControlEnabled('TCF_GENERAL', not TousCpt.Checked);
  SetControlVisible('BDEFAIRE'   , not TousCpt.Checked);
  SetControlVisible('IDEV1'      , TousCpt.Checked);
  SetControlEnabled('LBDEVISE'   , TousCpt.Checked);
  SetControlEnabled('CBDEVISE'   , TousCpt.Checked);

  GestionDevise;

  if TousCpt.Checked then begin
    TFFiche(Ecran).OnCloseQuery     := FicheCloseQuery;
    TFFiche(Ecran).BValider.OnClick := MajEnSerie;
  end
  else begin
    TFFiche(Ecran).BValider.OnClick := MajUnSeul;
    TFFiche(Ecran).OnCloseQuery     := FOnCloseQuery;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  {Initialisation de TCF_SOCIETE avec le paramsoc}
  SetField('TCF_SOCIETE', GetParamSocSecur('SO_SOCIETE', '001'));

  {Duplication d'une condition}
  if DuplicationOk then begin
    Q := OpenSQL('SELECT * FROM CONDITIONFINPLAC WHERE TCF_GENERAL ="' + General + '" AND TCF_CONDITIONFP = "' + Condition + '"', True);
    try
      if not Q.EOF then begin
        SetField('TCF_CONDITIONFP'   , Q.FindField('TCF_CONDITIONFP'   ).AsString);
        SetField('TCF_AGIOSPRECOMPTE', Q.FindField('TCF_AGIOSPRECOMPTE').AsString);
        SetField('TCF_AGIOSDEDUIT'   , Q.FindField('TCF_AGIOSDEDUIT'   ).AsString);
        SetField('TCF_BASECALCUL'    , Q.FindField('TCF_BASECALCUL'    ).AsString);
        SetField('TCF_TAUXFIXE'      , Q.FindField('TCF_TAUXFIXE'      ).AsString);
        SetField('TCF_TAUXPRECOMPTE' , Q.FindField('TCF_TAUXPRECOMPTE' ).AsString);
        SetField('TCF_VALTAUX'       , Q.FindField('TCF_VALTAUX'       ).AsFloat);
        SetField('TCF_TAUXVAR'       , Q.FindField('TCF_TAUXVAR'       ).AsString);
        SetField('TCF_VALMAJORATION' , Q.FindField('TCF_VALMAJORATION' ).AsFloat);
        SetField('TCF_VALMULTIPLE'   , Q.FindField('TCF_VALMULTIPLE'   ).AsFloat);
        SetField('TCF_NBJOURBANQUE'  , Q.FindField('TCF_NBJOURBANQUE'  ).AsInteger);
        SetField('TCF_NBJOURENCAISS' , Q.FindField('TCF_NBJOURENCAISS' ).AsInteger);
        SetField('TCF_NBJOURMINAGIOS', Q.FindField('TCF_NBJOURMINAGIOS').AsInteger);
        SetField('TCF_DATEDEBUT'     , Q.FindField('TCF_DATEDEBUT'     ).AsDateTime);
        SetField('TCF_DATEFIN'       , Q.FindField('TCF_DATEFIN'       ).AsDateTime);
        SetField('TCF_MONTANT'       , Q.FindField('TCF_MONTANT'       ).AsFloat);
      end;
    finally
      Ferme(Q);
    end;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.MajEnSerie(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  Obj : TObjDevise;
  lFP : TStringList;
  MAJ : Boolean;

    {------------------------------------------------------------------------}
    procedure ConstituerEnreg(Compte, Dev : string);
    {------------------------------------------------------------------------}
    var
      Tx : Double;
      ch : string;
    begin
      Tx := Obj.GetParite(Devise, Dev);
      ch := GetControlText('TCF_VALTAUX');
      {L'enregistrement existe, on le met à jour}
      if lFP.IndexOf(Compte + ';' + GetControlText('TCF_CONDITIONFP')) >= 0 then begin
        SQL := 'UPDATE CONDITIONFINPLAC SET TCF_SOCIETE = "' + GetParamSocSecur('SO_SOCIETE', '001') + '", ';
        SQL := SQL + 'TCF_BANQUE = "' + GetControlText('TCF_BANQUE') + '", TCF_AGIOSPRECOMPTE = "' + GetControlText('TCF_AGIOSPRECOMPTE') + '", ';
        SQL := SQL + 'TCF_AGIOSDEDUIT = "' + GetControlText('TCF_AGIOSDEDUIT') + '", TCF_BASECALCUL = "' + GetControlText('TCF_BASECALCUL') + '", ';
        SQL := SQL + 'TCF_TAUXFIXE = "' + GetControlText('TCF_TAUXFIXE') + '", TCF_TAUXPRECOMPTE = "' + GetControlText('TCF_TAUXPRECOMPTE') + '", ';
        SQL := SQL + 'TCF_VALTAUX = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', TCF_TAUXVAR = "' + GetControlText('TCF_TAUXVAR') + '", ';
        SQL := SQL + 'TCF_VALMAJORATION = ' + StrFPoint(Valeur(GetControlText('TCF_VALMAJORATION'))) + ', TCF_VALMULTIPLE = ' + StrFPoint(Valeur(GetControlText('TCF_VALMULTIPLE'))) + ', ';
        SQL := SQL + 'TCF_NBJOURBANQUE = ' + GetControlText('TCF_NBJOURBANQUE') + ', TCF_NBJOURENCAISS = ' + GetControlText('TCF_NBJOURENCAISS') + ', ';
        SQL := SQL + 'TCF_NBJOURMINAGIOS = ' + GetControlText('TCF_NBJOURMINAGIOS') + ', TCF_DATEDEBUT = "' + UsDateTime(StrToDate(GetControlText('TCF_DATEDEBUT'))) + '", ';
        SQL := SQL + 'TCF_DATEFIN = "' + UsDateTime(StrToDate(GetControlText('TCF_DATEFIN'))) + '", TCF_MONTANT  = ' + StrFPoint(Tx * Valeur(GetControlText('TCF_MONTANT')));
        SQL := SQL + ' WHERE TCF_CONDITIONFP = "' + GetControlText('TCF_CONDITIONFP') + '" AND TCF_GENERAL = "' + Compte + '"';
      end
      {Sinon on le crée}
      else begin
        SQL := 'INSERT INTO CONDITIONFINPLAC (TCF_SOCIETE, TCF_BANQUE, TCF_CONDITIONFP, TCF_GENERAL, TCF_AGIOSPRECOMPTE, ';
        SQL := SQL + 'TCF_AGIOSDEDUIT, TCF_BASECALCUL, TCF_TAUXFIXE, TCF_TAUXPRECOMPTE, TCF_VALTAUX, ';
        SQL := SQL + 'TCF_TAUXVAR, TCF_VALMAJORATION, TCF_VALMULTIPLE, TCF_NBJOURBANQUE, TCF_NBJOURENCAISS, ';
        SQL := SQL + 'TCF_NBJOURMINAGIOS, TCF_DATEDEBUT, TCF_DATEFIN, TCF_MONTANT) ';
        SQL := SQL + 'VALUES ("' + GetParamSocSecur('SO_SOCIETE', '001') + '", "' + GetControlText('TCF_BANQUE') + '", "' + GetControlText('TCF_CONDITIONFP') + '", "';
        SQL := SQL + Compte + '", "' + GetControlText('TCF_AGIOSPRECOMPTE') + '", "' + GetControlText('TCF_AGIOSDEDUIT') + '", "';
        SQL := SQL + GetControlText('TCF_BASECALCUL') + '", "' + GetControlText('TCF_TAUXFIXE') + '", "';
        SQL := SQL + GetControlText('TCF_TAUXPRECOMPTE') + '", ';
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', "';
        SQL := SQL + GetControlText('TCF_TAUXVAR') + '", ' + StrFPoint(Valeur(GetControlText('TCF_VALMAJORATION'))) + ', ';
        SQL := SQL + StrFPoint(Valeur(GetControlText('TCF_VALMULTIPLE'))) + ', ' + GetControlText('TCF_NBJOURBANQUE') + ', ';
        SQL := SQL + GetControlText('TCF_NBJOURENCAISS') + ', ' + GetControlText('TCF_NBJOURMINAGIOS') + ', "';
        SQL := SQL + UsDateTime(StrToDate(GetControlText('TCF_DATEDEBUT'))) + '", "' + UsDateTime(StrToDate(GetControlText('TCF_DATEFIN'))) + '", ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCF_MONTANT'))) + ' )';
      end;
    end;

begin
  {JP 22/03/04 : Si LastError a été passé à un, il est impossible de fermer la fenêtre}
  LastError := 0;

  if THValComboBox(GetControl('CBDEVISE')).ItemIndex = -1 then begin
    LastError := 1; {JP 22/03/04 : c'est plus propre ainsi}
    HShowMessage('0;Création collective;Veuillez sélectionner une devise !;W;O;O;O;', '', '');
    SetFocusControl('CBDEVISE'); {JP 22/03/04 : c'est plus propre ainsi}
    Exit;
  end;

  if GetControlText('TCF_BANQUE') = '' then begin
    LastError := 1;
    HShowMessage('1;Création collective;Veuillez sélectionner une banque !;W;O;O;O;', '', '');
    SetFocusControl('TCF_BANQUE');
    Exit;
  end;

  if GetControlText('TCF_CONDITIONFP') = '' then begin
    LastError := 1;
    HShowMessage('1;Création collective;Veuillez sélectionner une Transaction !;W;O;O;O;', '', '');
    SetFocusControl('TCF_CONDITIONFP');
    Exit;
  end;

  Obj := TObjDevise.Create(V_PGI.DateEntree);
  lFP := TStringList.Create;
  try
    MAJ := False;
    {On regarde si des enregistrements existent déjà dans la table => on fera un update}
    Q := OpenSQL('SELECT TCF_GENERAL, TCF_CONDITIONFP FROM CONDITIONFINPLAC WHERE TCF_CONDITIONFP = "' +
                 GetControlText('TCF_CONDITIONFP') + '" AND TCF_BANQUE = "' + GetControlText('TCF_BANQUE') + '"', True);
    while not Q.EOF do begin
      lFP.Add(Q.Fields[0].AsString + ';' + Q.Fields[1].AsString);
      Q.Next;
    end;
    Ferme(Q);

    {Des conditions de financement existent déjà pour cette banque et cette transaction ...}
    if lFP.Count > 0 then
      {... On demande s'il faut les écraser}
      MAJ := HShowMessage('2;Création collective;Certaines conditions existent déjà.'#13'Voulez-vous les mettre à jour ?;' +
                          'Q;YN;Y;Y;', '', '') = mrYes;

    Q := OpenSQL('SELECT BQ_CODE, BQ_DEVISE FROM BANQUECP WHERE BQ_BANQUE = "' + GetControlText('TCF_BANQUE') + '"', True);
    while not Q.EOF do begin
      {Si la condition de financement n'existe pas déjà ou si une mise à jour a été demandée}
      if (lFP.IndexOf(Q.Fields[0].AsString + ';' + GetControlText('TCF_CONDITIONFP')) < 0) or MAJ then begin
        ConstituerEnreg(Q.Fields[0].AsString, Q.Fields[1].AsString);
        ExecuteSQL(SQL);
      end;
      Q.Next;
    end;
    Ferme(Q);
  finally
    FreeAndNil(Obj);
    FreeAndNil(lFP);
  end;

  Ecran.Close;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.MajUnSeul(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {On exécute le code hérité}
  FValiderClick(Sender);
  {Puis on ferme la fiche si les champs obligtoires sont renseignés}
  if LastError = 0 then Ecran.Close;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  GestionDevise;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.GestionDevise;
{---------------------------------------------------------------------------------------}
begin
  if TousCpt.Enabled and TousCpt.Checked then begin
    Devise := GetControlText('CBDEVISE');
    AssignDrapeau(TImage(GetControl('IDEV1')), Devise);
    SetControlVisible('IDEV', False);
    SetControlVisible('DEV', False);
  end
  else begin
    Devise := RetDeviseCompte(GetControlText('TCF_GENERAL'));
    AssignDrapeau(TImage(GetControl('IDEV')), Devise);
    SetControlCaption('DEV', Devise);
    SetControlVisible('IDEV', True);
    SetControlVisible('DEV', True);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONFINPLAC.BanqueOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {15/11/05 : FQ 10307 : En eAgl, le OnChange est exécuté au chargement. On n'exécute le code
              ci-dessous, que si on n'est pas en consulation}
  if not (DS.State in [dsEdit, dsInsert]) then Exit;

  if GetControlText('TCF_BANQUE') = '' then
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '')
  else begin
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '') +
                     ' AND BQ_BANQUE = "' + cbBanque.Value + '"';
    SetControlText('TCF_GENERAL', '');
  end;
end;

initialization
  RegisterClasses([TOM_CONDITIONFINPLAC]);

end.

