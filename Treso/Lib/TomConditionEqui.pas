{-------------------------------------------------------------------------------------
  Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               08/11/01   BT   Création de l'unité
               03/08/03   JP   Migrtion eAgl
 1.20.001.001  22/03/04   JP   Petites corrections dans MajEnSerie (LastError et devise)
 1.50.001.001  28/04/04   JP   Petites corrections dans le OnNewRecord (il y avait des problèmes
                               d'initialisation, notamment sur TCE_TYPESOLDE
 6.20.001.002  19/01/05   JP   Gestion d'une duplication digne de ce nom !!
 7.00.001.012  13/06/06   JP   FQ 10368 : dans ToleranceOnClick on passait en Edit même si on était en Insert
 7.05.001.001  23/10/06   JP   Gestion des filtres multi sociétés 
--------------------------------------------------------------------------------------}
unit TomConditionEqui ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTob, 
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Fiche, HDB,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, SysUtils, HCtrls, HEnt1, UTOM;

type
  TOM_CONDITIONEQUI = class (TOM)
    {$IFDEF EAGLCLIENT}
    edCompte : THEdit;
    cbAgence : THValComboBox;
    {$ELSE}
    edCompte : THDBEdit;
    cbAgence : THDBValComboBox;
    {$ENDIF EAGLCLIENT}
    procedure OnNewRecord           ; override;
    procedure OnLoadRecord          ; override;
    procedure OnArgument(S : string); override;
    procedure OnUpdateRecord        ; override;
    //procedure OnChangeField(F : TField);override;
  protected
    PasTouche   : Boolean;
    Compte      : string;
    rbTolerance : TRadioButton;
    rbSoldeCons : TRadioButton;
    TousCpt     : TCheckBox;

    FValiderClick : TNotifyEvent;
    FOnCloseQuery : TCloseQueryEvent;
    Devise        : string;
    General       : string;
    DuplicationOk : Boolean; {19/01/05}

    procedure MajEnSerie      (Sender : TObject);
    procedure MajUnSeul       (Sender : TObject);
    procedure FicheCloseQuery (Sender : TObject; var CanClose: Boolean);
    procedure AgenceOnChange  (Sender : TObject);
    procedure ToleranceOnClick(Sender : TObject);
    procedure TraiterParAgence(Sender : TObject);
    procedure DeviseOnChange  (Sender : TObject);
    procedure CompteOnChange  (Sender : TObject);
    procedure GestionDevise   ;
  end ;

procedure TRLanceFiche_ConditionEqui(Dom, Fiche, Range, Lequel, Arguments: string);

implementation

uses
  Constantes, Commun, ExtCtrls, HMsgBox, ParamSoc, UObjGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ConditionEqui(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
var
  N : Integer;
begin
  inherited;
  Ecran.HelpContext := 150;
  ReadTokenSt(S);
  General := ReadTokenSt(S);
  {19/01/05 : Mise en place d'une duplication digne de ce nom}
  DuplicationOk := ReadTokenSt(S) = CODEDUPLICAT;

  rbTolerance := TRadioButton(GetControl('RBTOLERANCE'));
  rbSoldeCons := TRadioButton(GetControl('RBSOLDECONS'));

  TousCpt := TCheckBox(GetControl('CKAGENCE'));
  TousCpt.OnClick := TraiterParAgence;
  {La case à cocher n'est visible qu'en création}
  TousCpt.Enabled := TFFiche(Ecran).TypeAction in [taCreat, taCreatEnSerie, taCreatOne];
  GestionDevise;
  FValiderClick := TFFiche(Ecran).BValider.OnClick;
  FOnCloseQuery := Ecran.OnCloseQuery;
  TFFiche(Ecran).BValider.OnClick := MajUnSeul;
  THValComboBox(GetControl('CBDEVISE')).OnChange := DeviseOnChange;

  {On ne peut modifier l'agence qu'en création, puisqu'en consultation on ne peut modifier le
   compte qui fait partie de l'index}
  SetControlEnabled('TCE_AGENCE', TousCpt.Enabled);

  rbTolerance.OnClick := ToleranceOnClick;
  rbSoldeCons.OnClick := ToleranceOnClick;
  {$IFDEF EAGLCLIENT}
  edCompte := THEdit(GetControl('TCE_GENERAL'));
  cbAgence := THValComboBox(GetControl('TCE_AGENCE'));
  {$ELSE}
  edCompte := THDBEdit(GetControl('TCE_GENERAL'));
  cbAgence := THDBValComboBox(GetControl('TCE_AGENCE'));
  {$ENDIF EAGLCLIENT}
  cbAgence.OnChange := AgenceOnChange;
  edCompte.OnChange := CompteOnChange;

  SetControlProperty('PAGES', 'TABWIDTH', GetControl('PAGES').Width - 4);

  PasTouche := False;

  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '');

  N := Pos(';', S);
  if N > 0 then // Appel depuis TofEquilibrage
    Compte := Copy(S, N + 1, 999);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/01/2002
Modifié le ... :   /  /
Description .. : Affiche tous les comptes de l'agence
Mots clefs ... :
*****************************************************************}
procedure TOM_CONDITIONEQUI.AgenceOnChange(Sender: TObject);
begin
  if GetControlText('TCE_AGENCE') = '' then
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '')
  else begin
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '') + 
                     ' AND BQ_AGENCE = "' + cbAgence.Value + '"';
    if DS.State in [dsEdit, dsInsert] then
      edCompte.Text := '';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.ToleranceOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  c : string;
begin
  SetControlEnabled('TCE_SOLDEDEB' , rbTolerance.Checked);
  SetControlEnabled('TCE_SOLDECRED', rbTolerance.Checked);
  SetControlEnabled('TCE_SOLDECONS', not rbTolerance.Checked);

  if rbTolerance.Checked then C := 'X'
                         else C := '-';

  if not PasTouche then begin
    {13/06/06 : FQ 10368 : ne pas passer en Edit si on est en Insert}
    if not (DS.State in [dsEdit, dsInsert]) then
      DS.Edit;	// Modif pour demande de sauvegarde en sortie
    SetField('TCE_TYPESOLDE', C);
    SetField('TCE_COMPTE', GetField('TCE_COMPTE'));
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  SetControlEnabled('TCE_AGENCE', True);
  {Initialisation de TCE_SOCIETE avec le paramsoc}
  SetField('TCE_SOCIETE', GetParamSocSecur('SO_SOCIETE', '001'));

  {19/01/05 : On vient du mul et on veut dupliquer le paramétrage de "General"}
  if DuplicationOk then begin
    Q := OpenSQL('SELECT * FROM CONDITIONEQUI WHERE TCE_GENERAL = "' + General +  '"', True);
    try
      if not Q.EOF then begin
        SetField('TCE_SOLDECRED', Q.FindField('TCE_SOLDECRED').AsFloat);
        SetField('TCE_SOLDECONS', Q.FindField('TCE_SOLDECONS').AsFloat);
        SetField('TCE_DEBITMIN' , Q.FindField('TCE_DEBITMIN' ).AsFloat);
        SetField('TCE_CREDITMIN', Q.FindField('TCE_CREDITMIN').AsFloat);
        SetField('TCE_TYPESOLDE', Q.FindField('TCE_TYPESOLDE').AsString);
        SetField('TCE_ARRONDI'  , Q.FindField('TCE_ARRONDI'  ).AsString);
        {Pour le ToleranceOnClick qui est exécuté uniquement lorsqu'un radio est passé à True (mais pas à False)}
        PasTouche := True;
        rbTolerance.Checked := GetField('TCE_TYPESOLDE') = 'X';
        rbSoldeCons.Checked := not rbTolerance.Checked;
        PasTouche := False;
      end;
    finally
      Ferme(Q);
    end;
  end

  else begin
    {28/04/04 : On vient de la fiche d'équilibrage, on connait le compte, on recherche l'agence}
    if General <> '' then begin
      Q := OpenSQL('SELECT BQ_AGENCE FROM BANQUECP WHERE BQ_CODE ="' + General + '"', True);
      if not Q.EOF then begin
        SetField('TCE_AGENCE' , Q.Fields[0].AsString);
        SetField('TCE_GENERAL' , General);
      end;
      Ferme(Q);
    end;

    {28/04/04 : Initialisation des champs}
    SetField('TCE_TYPESOLDE', 'X');
    SetField('TCE_ARRONDI', '0'); {Par défaut, pas d'arrondi}
    rbTolerance.Checked := True;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.OnLoadRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if Ds.State <> dsInsert then begin
    PasTouche := True; // Evite modif table dans ToleranceOnClick
    if GetField('TCE_TYPESOLDE') = 'X' then
      rbTolerance.Checked := True
    else
      rbSoldeCons.Checked := True;
    PasTouche := False;
  end;
  {06/08/04 : FQ 10043 afficher le libellé plutôt que le compte général}
  Ecran.Caption := 'Conditions d''équilibrage du compte ' + RechDom('TRBANQUECP', GetField('TCE_GENERAL'), False);
  UpdateCaption(Ecran);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (GetField('TCE_GENERAL') = '') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le compte bancaire.');
    SetFocusControl('TCE_GENERAL');
  end

  else if (GetField('TCE_AGENCE') = '') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner une agence.');
    SetFocusControl('TCE_AGENCE');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.MajEnSerie(Sender : TObject);
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
    begin
      Tx := Obj.GetParite(Devise, Dev);
      {L'enregistrement existe, on le met à jour}
      if lFP.IndexOf(Compte) >= 0 then begin
        SQL := 'UPDATE CONDITIONEQUI SET TCE_SOLDEDEB = ' + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDEDEB'))) + ', ';
        SQL := SQL + 'TCE_SOLDECRED = ' + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDECRED'))) + ', ';
        SQL := SQL + 'TCE_SOLDECONS = ' + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDECONS'))) + ', ';
        SQL := SQL + 'TCE_DEBITMIN = '  + StrFPoint(Tx * Valeur(GetControlText('TCE_DEBITMIN' ))) + ', ';
        SQL := SQL + 'TCE_CREDITMIN = ' + StrFPoint(Tx * Valeur(GetControlText('TCE_CREDITMIN'))) + ', ';
        SQL := SQL + 'TCE_ARRONDI = "' + GetControlText('TCE_ARRONDI') + '", TCE_TYPESOLDE = "' + GetControlText('TCE_TYPESOLDE') + '" ';
        SQL := SQL + 'WHERE TCE_GENERAL = "' + Compte + '"';
      end
      {Sinon on le crée}
      else begin
        SQL := 'INSERT INTO CONDITIONEQUI (TCE_SOCIETE, TCE_AGENCE, TCE_GENERAL, TCE_SOLDEDEB, ';
        SQL := SQL + 'TCE_SOLDECRED, TCE_SOLDECONS, TCE_DEBITMIN, TCE_CREDITMIN, TCE_ARRONDI, TCE_TYPESOLDE) ';
        SQL := SQL + 'VALUES ("' + GetParamSocSecur('SO_SOCIETE', '001') + '", "' + GetControlText('TCE_AGENCE') + '", "' + Compte + '", ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDEDEB'))) + ', ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDECRED'))) + ', ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCE_SOLDECONS'))) + ', ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCE_DEBITMIN'))) + ', ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCE_CREDITMIN'))) + ', "';
        SQL := SQL + GetControlText('TCE_ARRONDI') + '", "' + GetControlText('TCE_TYPESOLDE') + '" )';
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

  if GetControlText('TCE_AGENCE') = '' then begin
    LastError := 1;
    HShowMessage('2;Création collective;Veuillez sélectionner une agence !;W;O;O;O;', '', '');
    SetFocusControl('TCE_AGENCE');
    Exit;
  end;

  Obj := TObjDevise.Create(V_PGI.DateEntree);
  lFP := TStringList.Create;
  try
    MAJ := False;
    {On regarde si des enregistrements existent déjà dans la table => on fera un update}
    Q := OpenSQL('SELECT TCE_GENERAL FROM CONDITIONEQUI WHERE TCE_AGENCE = "' + GetControlText('TCE_AGENCE') + '"', True);
    while not Q.EOF do begin
      lFP.Add(Q.Fields[0].AsString);
      Q.Next;
    end;
    Ferme(Q);

    {Des conditions d'équilibrage existent déjà pour cette banque et cette transaction ...}
    if lFP.Count > 0 then
      {... On demande s'il faut les écraser}
      MAJ := HShowMessage('2;Création collective;Certaines conditions existent déjà.'#13'Voulez-vous les mettre à jour ?;' +
                          'Q;YN;Y;Y;', '', '') = mrYes;

    Q := OpenSQL('SELECT BQ_CODE, BQ_DEVISE FROM BANQUECP WHERE BQ_AGENCE = "' + GetControlText('TCE_AGENCE') + '"', True);
    while not Q.EOF do begin
      {Si la condition d'équilibrage n'existe pas déjà ou si une mise à jour a été demandée}
      if (lFP.IndexOf(Q.Fields[0].AsString) < 0) or MAJ then begin
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
procedure TOM_CONDITIONEQUI.MajUnSeul(Sender : TObject);
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
procedure TOM_CONDITIONEQUI.FicheCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := LastError = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.TraiterParAgence(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not TousCpt.Enabled then Exit;
  SetControlEnabled('TCE_GENERAL', not TousCpt.Checked);
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
procedure TOM_CONDITIONEQUI.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  GestionDevise;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.GestionDevise;
{---------------------------------------------------------------------------------------}
begin
  if TousCpt.Enabled and TousCpt.Checked then begin
    Devise := GetControlText('CBDEVISE');
    AssignDrapeau(TImage(GetControl('IDEV1')), Devise);
    SetControlVisible('IDEV', False);
    SetControlVisible('DEV', False);
  end
  else begin
    Devise := RetDeviseCompte(GetControlText('TCE_GENERAL'));
    AssignDrapeau(TImage(GetControl('IDEV')), Devise);
    SetControlCaption('DEV', Devise);
    SetControlVisible('IDEV', True);
    SetControlVisible('DEV', True);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONEQUI.CompteOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {Maj des zones de gestions des devises}
  GestionDevise;
end;

initialization
  RegisterClasses([TOM_CONDITIONEQUI]);

end.

