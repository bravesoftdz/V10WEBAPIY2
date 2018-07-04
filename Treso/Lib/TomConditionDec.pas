{-------------------------------------------------------------------------------------
  Version   |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               05/12/01  BT   Création de l'unité
               03/08/03  JP   Migration en eAGL
    1.00       27/01/04  JP   Supression des références à TCN_PLAFONNEE : cf TCN_PLAFONNEE
 1.20.001.001  22/03/04  JP   Petites corrections dans MajEnSerie (LastError et devise)
 6.0X.001.001  19/10/04  JP   Gestion du plafond pour la CPFD
 6.20.001.002  19/01/05  JP   Gestion d'une duplication digne de ce nom !!
 7.09.001.001  23/10/06  JP   Gestion des filtres multi sociétés
 7.09.001.004  01/02/07  JP   FQ 10403 : Plantage à la validation car le champ Plafond est null
--------------------------------------------------------------------------------------}
unit TomConditionDec ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTOB,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Fiche, Hdb,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, SysUtils, ExtCtrls, HCtrls, HEnt1,
  UTOM;

type
  TOM_CONDITIONDEC = Class (TOM)
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
  private
    TousCpt : TCheckBox;

    FValiderClick : TNotifyEvent;
    FOnCloseQuery : TCloseQueryEvent;
    Devise        : string;
    General       : string;
    DuplicationOk : Boolean; {19/01/05}

    procedure MajEnSerie      (Sender : TObject);
    procedure MajUnSeul       (Sender : TObject);
    procedure FicheCloseQuery (Sender : TObject; var CanClose: Boolean);
    procedure TraiterParAgence(Sender : TObject);
    procedure DeviseOnChange  (Sender : TObject);
    procedure CompteOnChange  (Sender : TObject);
    procedure AgenceOnChange  (Sender : TObject);
    procedure GestionDevise   ;
  end ;

procedure TRLanceFiche_ConditionDec(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF VER150} variants,{$ENDIF}
  Constantes, Htb97, HMsgBox, Commun, ParamSoc, UObjGen, UProcGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_ConditionDec(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
var
  S : string;
begin
  S := Lequel;

  {Afin d'éviter le message "Enregistrement inaccessible", lorsque Lequel est rensegné}
  if S <> '' then
    if not ExisteSQL('SELECT TCN_FLUX FROM CONDITIONDEC WHERE TCN_FLUX = "' +
                     ReadTokenSt(S) + '" AND TCN_GENERAL = "' + ReadTokenSt(S) +  '"') then begin
      Arguments := 'ACTION=CREATION';
    end;
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  {Héritage de Maison Alfort qui est bien inutile aujourd'hui : à supprimer lors de la prochaine modif du MCD,
   sauf si l'on utilise la même table pour le paramétrage de comptes rémunérés}
  SetField('TCN_FLUX', 'DEC');
  {Initialisation de TCN_SOCIETE avec le paramsoc}
  SetField('TCN_SOCIETE', V_PGI.CodeSociete);

  {19/01/05 : On vient du mul et on veut dupliquer le paramétrage de "General"}
  if DuplicationOk then begin
    Q := OpenSQL('SELECT * FROM CONDITIONDEC WHERE TCN_FLUX = "DEC" AND TCN_GENERAL = "' + General +  '"', True);
    try
      if not Q.EOF then begin
        SetField('TCN_PLAFONDCPFD'  , Q.FindField('TCN_PLAFONDCPFD'  ).AsFloat);
        SetField('TCN_CPFD'         , Q.FindField('TCN_CPFD'         ).AsFloat);
        SetField('TCN_AUTORISATION' , Q.FindField('TCN_AUTORISATION' ).AsFloat);
        SetField('TCN_MAJOTAUX1'    , Q.FindField('TCN_MAJOTAUX1'    ).AsFloat);
        SetField('TCN_MAJOTAUX2'    , Q.FindField('TCN_MAJOTAUX2'    ).AsFloat);
        SetField('TCN_MAJOTAUX3'    , Q.FindField('TCN_MAJOTAUX3'    ).AsFloat);
        SetField('TCN_MULTAUX1'     , Q.FindField('TCN_MULTAUX1'     ).AsFloat);
        SetField('TCN_MULTAUX2'     , Q.FindField('TCN_MULTAUX2'     ).AsFloat);
        SetField('TCN_MULTAUX3'     , Q.FindField('TCN_MULTAUX3'     ).AsFloat);
        SetField('TCN_PLAFOND1'     , Q.FindField('TCN_PLAFOND1'     ).AsFloat);
        SetField('TCN_PLAFOND2'     , Q.FindField('TCN_PLAFOND2'     ).AsFloat);
        SetField('TCN_BASECALCUL'   , Q.FindField('TCN_BASECALCUL'   ).AsString);
        SetField('TCN_LIEAUTO'      , Q.FindField('TCN_LIEAUTO'      ).AsString);
        SetField('TCN_TYPECALCFRAIS', Q.FindField('TCN_TYPECALCFRAIS').AsString);
        SetField('TCN_TAUXREF1'     , Q.FindField('TCN_TAUXREF1'     ).AsString);
        SetField('TCN_TAUXREF2'     , Q.FindField('TCN_TAUXREF2'     ).AsString);
        SetField('TCN_TAUXREF3'     , Q.FindField('TCN_TAUXREF3'     ).AsString);
        SetField('TCN_PERIODE'      , Q.FindField('TCN_PERIODE'      ).AsString);
        SetField('TCN_CALCULSOLDE'  , Q.FindField('TCN_CALCULSOLDE'  ).AsString);
        SetField('TCN_SOLDEVALEUR'  , Q.FindField('TCN_SOLDEVALEUR'  ).AsString);
        SetField('TCN_TYPEPREMIER'  , Q.FindField('TCN_TYPEPREMIER'  ).AsString);
        SetField('TCN_NBJOUR'       , Q.FindField('TCN_NBJOUR'       ).AsInteger);
        SetField('TCN_DATECONTRAT'  , Q.FindField('TCN_DATECONTRAT'  ).AsDateTime);
      end;
    finally
      Ferme(Q);
    end;
  end

  {On n'est pas en duplication}
  else begin
    {Taux par défaut, généralement pratiqué par les banques}
    SetField('TCN_CPFD', Valeur(StrFPoint(0.0500)));
    {19/10/04 : En général, le plafonnement est de 50% des intérêts débiteurs}
    SetField('TCN_PLAFONDCPFD', Valeur(StrFPoint(50.00)));
    {19/01/05 : c'est plus propre}
    SetField('TCN_LIEAUTO', '-');

    {Si on vient des échelles d'intérêts, Général est renseigné}
    if General <> '' then begin
      SetField('TCN_GENERAL', General);
      GestionDevise;
      Q := OpenSQL('SELECT BQ_AGENCE FROM BANQUECP WHERE BQ_CODE = "' + General + '"', True);
      if not Q.EOF then
        SetField('TCN_AGENCE', Q.Fields[0].AsString);
      Ferme(Q);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.OnLoadRecord ;
{---------------------------------------------------------------------------------------}
var
  str        : string ;
begin
  inherited ;
  str := str + 'Standard : Tous mouvements ayant une date de valeur' ;
  str := str + #13 ;
  str := str + '           comprise dans la fourchette de sélection' ;
  str := str + #13 ;
  str := str + '           sont pris en compte.' ;
  str := str + #13 ;
  str := str + 'Bancaire : Tous mouvements sauf ceux ayant une date' ;
  str := str + #13 ;
  str := str + '           d''opération est supérieure à la fourchette';
  str := str + #13 ;
  str := str + '           de sélection sont pris en compte.' ;
  THEdit(GetControl('TCN_SOLDEVALEUR')).Hint := str ;
  {06/08/04 : FQ 10043 afficher le libellé plutôt que le compte général}
  Ecran.Caption := 'Conditions de découvert du compte ' + RechDom('TRBANQUECP', GetField('TCN_GENERAL'), False);
  UpdateCaption(Ecran);
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
var
  Ch : string;
begin
  inherited;
  Ecran.HelpContext := 500012;
  DuplicationOk := False;

  Ch := TFFiche(Ecran).FLequel;
  ReadTokenSt(Ch);
  General := ReadTokenSt(Ch);

  {19/01/05 : Mise en place d'une duplication digne de ce nom}
  if General = '' then begin
    ReadTokenSt(S);
    General := ReadTokenSt(S);
    DuplicationOk := ReadTokenSt(S) = CODEDUPLICAT;
  end;

  TousCpt := TCheckBox(GetControl('CKAGENCE'));
  TousCpt.OnClick := TraiterParAgence;
  {La case à cocher n'est visible qu'en création}
  TousCpt.Enabled := TFFiche(Ecran).TypeAction in [taCreat, taCreatEnSerie, taCreatOne];
  {On ne peut modifier l'agence qu'en création, puisqu'en consultation on ne peut modifier le
   compte qui fait partie de l'index}
  SetControlEnabled('TCE_AGENCE', TousCpt.Enabled);

  GestionDevise;

  FValiderClick := TFFiche(Ecran).BValider.OnClick;
  FOnCloseQuery := Ecran.OnCloseQuery;
  TFFiche(Ecran).BValider.OnClick := MajUnSeul;

  THValComboBox(GetControl('CBDEVISE')).OnChange := DeviseOnChange;
  {$IFDEF EAGLCLIENT}
  edCompte := THEdit(GetControl('TCN_GENERAL'));
  cbAgence := THValComboBox(GetControl('TCN_AGENCE'));
  {$ELSE}
  edCompte := THDBEdit(GetControl('TCN_GENERAL'));
  cbAgence := THDBValComboBox(GetControl('TCN_AGENCE'));
  {$ENDIF EAGLCLIENT}
  edCompte.OnChange := CompteOnChange;
  cbAgence.OnChange := AgenceOnChange;
  SetPlusBancaire(cbAgence, 'TRA', CODECOURANTS + ';' + CODETITRES + ';');
  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {02/01/07 : FQ 10403 : Valeur(GetField('TCN_PLAFONDCPFD')) plante si null, j'initialise le champ}
  if VarToStr(GetField('TCN_PLAFONDCPFD')) = '' then SetField('TCN_PLAFONDCPFD', 0);
  
  if (GetField('TCN_GENERAL') = '') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le compte bancaire.');
    SetFocusControl('TCN_GENERAL');
  end

  else if GetControlText('TCN_AGENCE') = '' then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Veuillez sélectionner une agence.');
    SetFocusControl('TCN_AGENCE');
  end

  else if not Between(Valeur(GetField('TCN_PLAFONDCPFD')), 0, 100) then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Le plafonnement du plus fort découvert doit être compris entre 0 et 100%.');
    SetFocusControl('TCN_PLAFONDCPFD');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.MajEnSerie(Sender : TObject);
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
      {L'enregistrement existe, on le met à jour}
      if lFP.IndexOf(Compte) >= 0 then begin
        SQL := 'UPDATE CONDITIONDEC SET TCN_AUTORISATION = ' + StrFPoint(Tx * Valeur(GetControlText('TCN_AUTORISATION'))) + ', ';
        SQL := SQL + 'TCN_TYPECALCFRAIS = "' + GetControlText('TCN_TYPECALCFRAIS') + '", ';
        ch := GetControlText('TCN_MAJOTAUX1');
        SQL := SQL + 'TCN_MAJOTAUX1 =' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MAJOTAUX2');
        SQL := SQL + 'TCN_MAJOTAUX2 = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MAJOTAUX3');
        SQL := SQL + 'TCN_MAJOTAUX3 = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX1');
        SQL := SQL + 'TCN_MULTAUX1 = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX2');
        SQL := SQL + 'TCN_MULTAUX2 = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX3');
        SQL := SQL + 'TCN_MULTAUX3 = ' + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        SQL := SQL + 'TCN_TAUXREF1 = "' + GetControlText('TCN_TAUXREF1') + '", TCN_TAUXREF3 = "' + GetControlText('TCN_TAUXREF3') + '", ';
        SQL := SQL + 'TCN_TAUXREF2 = "' + GetControlText('TCN_TAUXREF2') + '", TCN_PLAFOND1 = ' + StrFPoint(Tx * Valeur(GetControlText('TCN_PLAFOND1'))) + ', ';
        SQL := SQL + 'TCN_PLAFOND2 = ' + StrFPoint(Tx * Valeur(GetControlText('TCN_PLAFOND2'))) + ', ';
        SQL := SQL + 'TCN_BASECALCUL = "' + GetControlText('TCN_BASECALCUL') + '", TCN_DATECONTRAT = "' + UsDateTime(StrToDate(GetControlText('TCN_DATECONTRAT'))) + '", ';
        SQL := SQL + 'TCN_PERIODE = "' +GetControlText('TCN_PERIODE') + '", TCN_CALCULSOLDE = "' + GetControlText('TCN_CALCULSOLDE') + '", ';
        SQL := SQL + 'TCN_SOLDEVALEUR = "' + GetControlText('TCN_SOLDEVALEUR') + '", TCN_NBJOUR = ' + GetControlText('TCN_NBJOUR') + ', ';
        SQL := SQL + 'TCN_TYPEPREMIER = "' + GetControlText('TCN_TYPEPREMIER') + '", TCN_CPFD = ';
        ch := GetControlText('TCN_CPFD');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        SQL := SQL + {'TCN_PLAFONNEE = "' + GetControlText('TCN_PLAFONNEE') + '",}' TCN_LIEAUTO = "' + GetControlText('TCN_LIEAUTO') + '" ';
        SQL := SQL + 'WHERE TCN_GENERAL = "' + Compte + '"';
      end
      {Sinon on le crée}
      else begin
        SQL := 'INSERT INTO CONDITIONDEC (TCN_SOCIETE, TCN_AGENCE, TCN_GENERAL, TCN_FLUX, TCN_AUTORISATION, ';
        SQL := SQL + 'TCN_TYPECALCFRAIS, TCN_MAJOTAUX1, TCN_MAJOTAUX2, TCN_MAJOTAUX3, TCN_MULTAUX1, TCN_MULTAUX2, ';
        SQL := SQL + 'TCN_MULTAUX3, TCN_TAUXREF1, TCN_TAUXREF2, TCN_TAUXREF3, TCN_PLAFOND1, TCN_PLAFOND2, ';
        SQL := SQL + 'TCN_BASECALCUL, TCN_DATECONTRAT, TCN_PERIODE, TCN_CALCULSOLDE, TCN_SOLDEVALEUR, TCN_NBJOUR, ';
        SQL := SQL + 'TCN_TYPEPREMIER, TCN_CPFD, ' +{TCN_PLAFONNEE,} 'TCN_LIEAUTO) ';
        SQL := SQL + 'VALUES ("' + GetParamSocSecur('SO_SOCIETE', '001') + '", "' + GetControlText('TCN_AGENCE') + '", "';
        SQL := SQL + Compte + '", "DEC", ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCN_AUTORISATION'))) + ', "';
        SQL := SQL + GetControlText('TCN_TYPECALCFRAIS') + '", ';
        ch := GetControlText('TCN_MAJOTAUX1');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MAJOTAUX2');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MAJOTAUX3');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX1');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX2');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', ';
        ch := GetControlText('TCN_MULTAUX3');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', "' + GetControlText('TCN_TAUXREF1') + '", "';
        SQL := SQL + GetControlText('TCN_TAUXREF2') + '", "' + GetControlText('TCN_TAUXREF3') + '", ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCN_PLAFOND1'))) + ', ';
        SQL := SQL + StrFPoint(Tx * Valeur(GetControlText('TCN_PLAFOND2'))) + ', "';
        SQL := SQL + GetControlText('TCN_BASECALCUL') + '", "' + UsDateTime(StrToDate(GetControlText('TCN_DATECONTRAT'))) + '", "';
        SQL := SQL + GetControlText('TCN_PERIODE') + '", "' + GetControlText('TCN_CALCULSOLDE') + '", "';
        SQL := SQL + GetControlText('TCN_SOLDEVALEUR') + '", ' + GetControlText('TCN_NBJOUR') + ', "';
        SQL := SQL + GetControlText('TCN_TYPEPREMIER') + '", ';
        ch := GetControlText('TCN_CPFD');
        SQL := SQL + StrFPoint(Valeur(ReadTokenPipe(ch, '%'))) + ', "';
        SQL := SQL + {GetControlText('TCN_PLAFONNEE') + '", "' +} GetControlText('TCN_LIEAUTO') + '" )';
      end;
    end;

begin
  {JP 22/03/04 : Si LastError a été passé à un, il est impossible de fermer la fenêtre}
  LastError := 0;

  if THValComboBox(GetControl('CBDEVISE')).ItemIndex = -1 then begin
    LastError := 1;
    HShowMessage('0;Création collective;Veuillez sélectionner une devise !;W;O;O;O;', '', '');
    SetFocusControl('CBDEVISE');
    Exit;
  end;

  if GetControlText('TCN_AGENCE') = '' then begin
    LastError := 1;
    HShowMessage('2;Création collective;Veuillez sélectionner une agence !;W;O;O;O;', '', '');
    SetFocusControl('TCN_AGENCE');
    Exit;
  end;

  Obj := TObjDevise.Create(V_PGI.DateEntree);
  lFP := TStringList.Create;
  try
    MAJ := False;
    {On regarde si des enregistrements existent déjà dans la table => on fera un update}
    Q := OpenSQL('SELECT TCN_GENERAL FROM CONDITIONDEC WHERE TCN_AGENCE = "' + GetControlText('TCN_AGENCE') + '"', True);
    while not Q.EOF do begin
      lFP.Add(Q.Fields[0].AsString);
      Q.Next;
    end;
    Ferme(Q);

    {Des conditions de découvert existent déjà pour cette banque et cette transaction ...}
    if lFP.Count > 0 then
      {... On demande s'il faut les écraser}
      MAJ := HShowMessage('2;Création collective;Certaines conditions existent déjà.'#13'Voulez-vous les mettre à jour ?;' +
                          'Q;YN;Y;Y;', '', '') = mrYes;

    Q := OpenSQL('SELECT BQ_CODE, BQ_DEVISE FROM BANQUECP WHERE BQ_AGENCE = "' + GetControlText('TCN_AGENCE') + '"', True);
    while not Q.EOF do begin
      {Si la condition de découvert n'existe pas déjà ou si une mise à jour a été demandée}
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
procedure TOM_CONDITIONDEC.MajUnSeul(Sender : TObject);
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
procedure TOM_CONDITIONDEC.FicheCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := LastError = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.TraiterParAgence(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not TousCpt.Enabled then Exit;
  SetControlEnabled('TCN_GENERAL', not TousCpt.Checked);
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
procedure TOM_CONDITIONDEC.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  GestionDevise;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.GestionDevise;
{---------------------------------------------------------------------------------------}
begin
  if TousCpt.Enabled and TousCpt.Checked then begin
    Devise := GetControlText('CBDEVISE');
    AssignDrapeau(TImage(GetControl('IDEV1')), Devise);
    SetControlVisible('IDEV', False);
    SetControlVisible('DEV', False);
  end
  else begin
    Devise := RetDeviseCompte(GetControlText('TCN_GENERAL'));
    AssignDrapeau(TImage(GetControl('IDEV')), Devise);
    SetControlCaption('DEV', Devise);
    SetControlVisible('IDEV', True);
    SetControlVisible('DEV', True);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.CompteOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  GestionDevise;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CONDITIONDEC.AgenceOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('TCN_AGENCE') = '' then
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '')
  else begin
    {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
    edCompte.Plus := FiltreBanqueCp(edCompte.DataType, '', '') +
                     'AND BQ_AGENCE = "' + cbAgence.Value + '"';
    SetControlText('TCN_GENERAL', '');
  end;
end;

initialization
  RegisterClasses([TOM_CONDITIONDEC]);

end.

