{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               25/10/01   BT   Création de l'unité
               03/08/03   JP   Migration en eAGL
 1.20.001.001  22/03/04   JP   Petites corrections dans MajEnSerie (LastError)
 1.90.xxx.xxx  18/06/04   JP   FQ10096 : Amélioration de la gestion du Plus sur les CIB
 6.20.001.002  19/01/05   JP   Gestion d'une duplication digne de ce nom !!
 6.51.001.001  15/11/05   JP   FQ 10307 : On court-circuite le onChange si en mode consultation
 7.05.001.001  23/10/06   JP   Gestion des filtres multi sociétés
--------------------------------------------------------------------------------------}
unit TomConditionVal ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTob,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Fiche, HDB,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, SysUtils, HCtrls, HEnt1, UTOM,
  HMsgBox, ParamSoc;

type
  TOM_CONDITIONVAL = class (TOM)
    procedure OnNewRecord               ; override;
    procedure OnArgument    (S : string); override;
    procedure OnUpdateRecord            ; override;
    procedure OnLoadRecord              ; override;
  private
    TousCpt       : TCheckBox;

    {$IFDEF EAGLCLIENT}
    edGeneral : THEdit;
    edCib     : THEdit;
    cbBanque  : THValComboBox;
    {$ELSE}
    edGeneral : THDBEdit;
    edCib     : THDBEdit;
    cbBanque  : THDBValComboBox;
    {$ENDIF}
    FValiderClick : TNotifyEvent;
    FOnCloseQuery : TCloseQueryEvent;
    General       : string;
    CodeCib       : string;
    DuplicationOk : Boolean; {19/01/05}

    procedure MajEnSerie      (Sender : TObject);
    procedure MajUnSeul       (Sender : TObject);
    procedure FicheCloseQuery (Sender : TObject; var CanClose: Boolean);
    procedure TraiterParAgence(Sender : TObject);
    procedure BanqueOnChange  (Sender : TObject);
  end ;

procedure TRLanceFiche_ConditionVal(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  Constantes, commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ConditionVal(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  ReadTokenSt(S);
  General := ReadTokenSt(S);
  CodeCib := ReadTokenSt(S);
  {19/01/05 : Correction de la duplication}
  DuplicationOk := ReadTokenSt(S) = CODEDUPLICAT;

  TousCpt := TCheckBox(GetControl('CKAGENCE'));
  TousCpt.OnClick := TraiterParAgence;

  {La case à cocher n'est active qu'en création}
  TGroupBox(GetControl('GBCOLLECT')).Enabled := TFFiche(Ecran).TypeAction in [taCreat, taCreatEnSerie, taCreatOne];
  TousCpt.Enabled := TGroupBox(GetControl('GBCOLLECT')).Enabled;

  FValiderClick := TFFiche(Ecran).BValider.OnClick;
  FOnCloseQuery := Ecran.OnCloseQuery;
  TFFiche(Ecran).BValider.OnClick := MajUnSeul;

  {On ne peut modifier la banque qu'en création, puisqu'en consultation on ne peut modifier le
   compte qui fait partie de l'index}
  SetControlEnabled('TCV_BANQUE', TousCpt.Enabled);

  {$IFDEF EAGLCLIENT}
  edGeneral := THEdit(GetControl('TCV_GENERAL'));
  edCib     := THEdit(GetControl('TCV_CODECIB'));
  cbBanque  := THValComboBox(GetControl('TCV_BANQUE'));
  {$ELSE}
  edGeneral := THDBEdit(GetControl('TCV_GENERAL'));
  edCib     := THDBEdit(GetControl('TCV_CODECIB'));
  cbBanque  := THDBValComboBox(GetControl('TCV_BANQUE'));
  {$ENDIF}

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  edGeneral.Plus := FiltreBanqueCp(edGeneral.DataType, '', '');

  SetControlProperty('PAGES', 'TABWIDTH', GetControl('PAGES').Width - 4);

  cbBanque.OnChange := BanqueOnChange;
  edCib.Plus := 'TCI_BANQUE = "' + CODECIBREF + '"';
  SetPlusBancaire(cbBanque, 'PQ', CODECOURANTS + ';' + CODETITRES + ';');
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.Caption := 'Conditions de valeur du compte ' + RechDom('TRBANQUECP', GetField('TCV_GENERAL'), False);
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (GetField('TCV_BANQUE') = '') then begin
    LastError := 1;
    HShowMessage('0;Conditions de valeur; Veuillez renseigner la banque !;W;O;O;O;', '', '');
    SetFocusControl('TCV_BANQUE');
    Exit;
  end;

  if (GetField('TCV_CODECIB') = '') then begin
    LastError := 1;
    HShowMessage('1;Conditions de valeur; Veuillez renseigner le code CIB !;W;O;O;O;', '', '');
    SetFocusControl('TCV_CODECIB');
    Exit;
  end;

  if (GetField('TCV_GENERAL') = '') then begin
    LastError := 1;
    HShowMessage('2;Conditions de valeur; Veuillez renseigner le compte bancaire !;W;O;O;O;', '', '');
    SetFocusControl('TCV_GENERAL');
    Exit;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
   
  {---------------------------------------------------------------}
  procedure InitCombo(cbName: string);
  {---------------------------------------------------------------}
  var	cbType : THDBValComboBox;
  begin
    {$IFDEF EAGLCLIENT}
    cbType := THValComboBox(GetControl(cbName));
    cbType.Value := '0';
    SetField(cbName, '0');

    {$ELSE}
    cbType := THDBValComboBox(GetControl(cbName));
    cbType.Field.AsString := '0';
    {$ENDIF}
  end;

begin
  inherited;
  {Initialisation de TCV_SOCIETE avec le paramsoc}
  SetField('TCV_SOCIETE', GetParamSocSecur('SO_SOCIETE', '001'));

  {Duplication d'une condition}
  if DuplicationOk then begin
    {Recherche de l'enregistrement source}
    Q := OpenSQL('SELECT * FROM CONDITIONVAL WHERE TCV_GENERAL ="' + General + '" AND TCV_CODECIB = "' + CodeCib + '"', True);
    try
      if not Q.EOF then begin
        SetField('TCV_CODECIB'       , Q.FindField('TCV_CODECIB'       ).AsString);
        SetField('TCV_TYPEGLISSEMENT', Q.FindField('TCV_TYPEGLISSEMENT').AsString);
        SetField('TCV_NBJPREMIER'    , Q.FindField('TCV_NBJPREMIER'    ).AsInteger);
        SetField('TCV_TYPEPREMIER'   , Q.FindField('TCV_TYPEPREMIER'   ).AsString);
        SetField('TCV_COMMISSION1'   , Q.FindField('TCV_COMMISSION1'   ).AsString);
        SetField('TCV_COMMISSION2'   , Q.FindField('TCV_COMMISSION2'   ).AsString);
        SetField('TCV_COMMISSION3'   , Q.FindField('TCV_COMMISSION3'   ).AsString);
        SetField('TCV_NBJDEUXIEME'   , Q.FindField('TCV_NBJDEUXIEME'   ).AsInteger);
        SetField('TCV_TYPEDEUXIEME'  , Q.FindField('TCV_TYPEDEUXIEME'  ).AsString);
      end;
    finally
      Ferme(Q);
    end;
  end
  {Création simple}
  else begin
    {Initialisation des combos des dates de valeur}
    InitCombo('TCV_TYPEGLISSEMENT');
    InitCombo('TCV_TYPEPREMIER');
    InitCombo('TCV_TYPEDEUXIEME');
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.MajEnSerie(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  SQL : string;
  lFP : TStringList;
  MAJ : Boolean;

    {------------------------------------------------------------------------}
    procedure ConstituerEnreg(Compte : string);
    {------------------------------------------------------------------------}
    begin
      {L'enregistrement existe, on le met à jour}
      if lFP.IndexOf(Compte + ';' + GetControlText('TCV_CODECIB')) >= 0 then begin
        SQL := 'UPDATE CONDITIONVAL SET TCV_TYPEGLISSEMENT = "' + GetControlText('TCV_TYPEGLISSEMENT') + '", ';
        SQL := SQL + 'TCV_NBJPREMIER = ' + GetControlText('TCV_NBJPREMIER') + ', TCV_TYPEPREMIER = "' + GetControlText('TCV_TYPEPREMIER') + '", ';
        SQL := SQL + 'TCV_COMMISSION1 = "' + GetControlText('TCV_COMMISSION1') + '", TCV_COMMISSION2 = "' + GetControlText('TCV_COMMISSION2') + '", ';
        SQL := SQL + 'TCV_COMMISSION3 = "' +  GetControlText('TCV_COMMISSION3') + '", TCV_NBJDEUXIEME = ' + GetControlText('TCV_NBJDEUXIEME') + ', ';
        SQL := SQL + 'TCV_TYPEDEUXIEME = "' + GetControlText('TCV_TYPEDEUXIEME') + '" ';
        SQL := SQL + 'WHERE TCV_GENERAL = "' + Compte + '" AND TCV_CODECIB = "' + GetControlText('TCV_CODECIB') + '"';
      end
      {Sinon on le crée}
      else begin
        SQL := 'INSERT INTO CONDITIONVAL (TCV_SOCIETE, TCV_BANQUE, TCV_GENERAL, TCV_CODECIB, ';
        SQL := SQL + 'TCV_TYPEGLISSEMENT, TCV_NBJPREMIER, TCV_TYPEPREMIER, TCV_COMMISSION1, TCV_COMMISSION2, ';
        SQL := SQL + ' TCV_COMMISSION3, TCV_NBJDEUXIEME, TCV_TYPEDEUXIEME) ';
        SQL := SQL + 'VALUES ("' + GetParamSocSecur('SO_SOCIETE', '001') + '", "' + GetControlText('TCV_BANQUE') + '", "';
        SQL := SQL + Compte + '", "' + GetControlText('TCV_CODECIB') + '", "' + GetControlText('TCV_TYPEGLISSEMENT') + '", ';
        SQL := SQL + GetControlText('TCV_NBJPREMIER') + ', "' + GetControlText('TCV_TYPEPREMIER') + '", "';
        SQL := SQL + GetControlText('TCV_COMMISSION1') + '", "' + GetControlText('TCV_COMMISSION2') + '", "';
        SQL := SQL + GetControlText('TCV_COMMISSION3') + '", ' + GetControlText('TCV_NBJDEUXIEME') + ', "';
        SQL := SQL + GetControlText('TCV_TYPEDEUXIEME') + '" )';
      end;
    end;

begin
  {JP 22/03/04 : Si LastError a été passé à un, il est impossible de fermer la fenêtre}
  LastError := 0;

  if GetControlText('TCV_BANQUE') = '' then begin
    LastError := 1;
    HShowMessage('1;Création collective;Veuillez sélectionner une banque !;W;O;O;O;', '', '');
    SetFocusControl('TCV_BANQUE');
    Exit;
  end;

  if GetControlText('TCV_CODECIB') = '' then begin
    LastError := 1;
    HShowMessage('2;Création collective;Veuillez sélectionner un code CIB !;W;O;O;O;', '', '');
    SetFocusControl('TCV_CODECIB');
    Exit;
  end;

  lFP := TStringList.Create;
  try
    MAJ := False;
    {On regarde si des enregistrements existent déjà dans la table => on fera un update}
    Q := OpenSQL('SELECT TCV_GENERAL, TCV_CODECIB FROM CONDITIONVAL WHERE TCV_CODECIB = "' +
                 GetControlText('TCV_CODECIB')  + '" AND TCV_BANQUE = "' + GetControlText('TCV_BANQUE') + '"', True);
    while not Q.EOF do begin
      lFP.Add(Q.Fields[0].AsString + ';' + Q.Fields[1].AsString);
      Q.Next;
    end;
    Ferme(Q);

    {Des conditions de valeurs existent déjà pour cette banque et ce CIB ...}
    if lFP.Count > 0 then
      {... On demande s'il faut les écraser}
      MAJ := HShowMessage('2;Création collective;Certaines conditions existent déjà.'#13'Voulez-vous les mettre à jour ?;' +
                          'Q;YN;Y;Y;', '', '') = mrYes;

    Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_BANQUE = "' + GetControlText('TCV_BANQUE') + '"', True);
    while not Q.EOF do begin
      {Si la condition de valeur n'existe pas déjà ou si une mise à jour a été demandée}
      if (lFP.IndexOf(Q.Fields[0].AsString + ';' + GetControlText('TCV_CODECIB')) < 0) or MAJ then begin
        ConstituerEnreg(Q.Fields[0].AsString);
        ExecuteSQL(SQL);
      end;
      Q.Next;
    end;
    Ferme(Q);
  finally
    FreeAndNil(lFP);
  end;

  Ecran.Close;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.MajUnSeul(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {On exécute le code hérité}
  FValiderClick(Sender);
  {Puis on ferme la fiche si les champs obligtoires sont renseignés}
  if LastError = 0 then Ecran.Close;
end;

{Surcharge du FormCloseQuery lors d'une création en série pour éviter le message "Voulez-
 cous enregistrer vos modification
{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.FicheCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := LastError = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONVAL.TraiterParAgence(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not TGroupBox(GetControl('GBCOLLECT')).Enabled then Exit;
  SetControlEnabled('TCV_GENERAL', not TousCpt.Checked);
  SetControlVisible('BDEFAIRE'   , not TousCpt.Checked);

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
procedure TOM_CONDITIONVAL.BanqueOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {15/11/05 : FQ 10307 : En eAgl, le OnChange est exécuté au chargement. On n'exécute le code
              ci-dessous, que si on n'est pas en consulation}
  if not (DS.State in [dsEdit, dsInsert]) then Exit;

  if GetControlText('TCV_BANQUE') = '' then begin
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edGeneral.Plus := FiltreBanqueCp(edGeneral.DataType, '', '');
    edCib.Plus := 'TCI_BANQUE = "' + CODECIBREF + '"';
  end
  else begin
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edGeneral.Plus := FiltreBanqueCp(edGeneral.DataType, '', '') + 
                      ' AND BQ_BANQUE = "' + cbBanque.Value + '"';
    edCib.Plus := 'TCI_BANQUE = "' + cbBanque.Value + '"';
    {18/06/04 : FQ 10096, modification de la gestion du filtre}
    SetField('TCV_GENERAL', '');
    if not ExisteSQL('SELECT TCI_CODECIB FROM CIB WHERE TCI_BANQUE = "' + GetControlText('TCV_BANQUE') + '" AND ' +
                     'TCI_CODECIB = "' + GetControlText('TCV_CODECIB') + '"') then
      SetField('TCV_CODECIB', '');
  end;
end;

initialization
  RegisterClasses([TOM_CONDITIONVAL]);

end.

