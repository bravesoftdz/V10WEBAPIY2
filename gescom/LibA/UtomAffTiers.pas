{***********UNITE*************************************************
Auteur  ...... : A.B
Créé le ...... : 09/02/2005
Modifié le ... : 09/02/2005
Description .. : Liste des intervenants depuis affaire et tâche
Mots clefs ... : TOM;AFFTIERS
*****************************************************************}
unit UtomAffTiers;

interface

uses
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
{$IFNDEF EAGLSERVER}
  Fiche,
  FichList,
  FE_Main,
{$ENDIF EAGLSERVER}
{$ELSE EAGLCLIENT}
  eFiche,
  eFichList,
  MaineAGL,
{$ENDIF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
  SaisieList,
{$ENDIF EAGLSERVER}
  Utob,
  sysutils,
  HCtrls,
  HEnt1,
  UTOM,
  Windows,
  AffaireUtil,
  DicoBTP,
  EntGC,
  stdctrls,
  wcommuns,
  TiersUtil;

type
  TOM_AFFTIERS = class (TOM)
    procedure OnAfterDeleteRecord;              override;
    procedure OnAfterUpdateRecord;              override;
    procedure OnArgument (stArgument: string);  override;
    procedure OnCancelRecord;                   override;
    procedure OnClose;                          override;
    procedure OnDeleteRecord;                   override;
    procedure OnLoadRecord;                     override;
    procedure OnNewRecord;                      override;
    procedure OnUpdateRecord;                   override;
  private
    Action        : TActionFiche;
    //
    LeClient      : THEdit;
    //
    TypeInterv    : THDBValcomboBox;
    LeContact     : THEdit;
    LaRessource   : THEdit;
    LeTiers       : THEdit;
    TsContact     : THEdit;
    //
    LibAffaire    : THLabel;
    LibClient     : THLabel;
    //
    fStType         : String;
    fStTypeOrig     : String;
    fStAffaire      : String;
    fStAuxiliaire   : String;
    fStLibOrig      : String;
    fStTypeInterv   : String;
    fStAffaireEcole : String;
    fStRang         : string;
    fInNumOrig      : Integer;

{$IFNDEF EAGLSERVER}
    fbSaisieList  : Boolean;
    fbFicheListe  : boolean;
{$ENDIF EAGLSERVER}

    procedure LoadAffaire;

{$IFNDEF EAGLSERVER}
    procedure ClientClick(Sender: TObject);
    procedure ContactTouche(Sender: TObject; var Key: char);
    procedure Controlechamp(Champ, Valeur: String);

    procedure GetScreenObject;

    procedure LeContactOnElipsisClick (Sender: TObject);

    procedure RessourceClick(Sender: TObject);
    procedure RessourceExit(Sender: TObject);
    procedure TiersExit(Sender: TObject);
    procedure TousContactOnElipsisClick (Sender: TObject);
    procedure TypeIntervenantOnClick (Sender: TObject);

{$ENDIF EAGLSERVER}

  end;

{$IFNDEF EAGLSERVER}
procedure AFLanceFiche_AFIntervenant (Range, lequel, argument: string);
{$ENDIF EAGLSERVER}

implementation
uses AGLInitGC, HDB;
const
  TexteMessage: array [1..5] of string = (
    {1}'Type d''intervenant obligatoire',
    {2}'Code intervenant obligatoire',
    {3}'Inscrit à ',
    {4}'Inscrit à la session de formation',
    {5}'Ce contact est déjà affecté, voulez-vous l''enregistrer ?'
    );

{$IFNDEF EAGLSERVER}
procedure AFLanceFiche_AFIntervenant (Range, lequel, argument: string);
begin
  AGLLanceFiche ('AFF', 'AFINTERVENANT', Range, lequel, Argument);
end;
{$ENDIF EAGLSERVER}

procedure TOM_AFFTIERS.OnArgument (stArgument: string);
var
  {$IFDEF GIGI}
  TypeInter : THDBVALComboBox;
  {$ENDIF GIGI}
  Critere   : string;
  Champ     : string;
  Valeur    : string;
  x         : Integer;
begin
  inherited;

  GetScreenObject();

  {$IFNDEF EAGLSERVER}
  if ecran <> nil then
  begin
    fbSaisieList := ( ECran is TFSaisieList ); //GA_200807_AB-FQ15139-Edition affaire école
    fbFicheListe := ( ECran is TFFicheListe );
  end;
  {$ENDIF !EAGLSERVER}

  if ctxScot in V_PGI.PGIContexte then
  begin
    SetControlVisible ('AFT_PROPOINTERV', False);
    SetControlVisible ('TAFT_PROPOINTERV', False);
  end;

  fstTypeOrig     := '';

  if (stArgument = '') then exit;

  //Récupération des paramètres de lancements.....
  Critere := uppercase(Trim(ReadTokenSt(StArgument)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ  := Critere;
     //
     ControleChamp(Champ, Valeur);
     //
     Critere:= uppercase(Trim(ReadTokenSt(StArgument)));
  end;
  fInNumOrig := 0;

  //On charge l'affaire passée en paramètres !!!
  LoadAffaire;

  UpdateCaption (Ecran);

  // mcd 02/11/2006 transtypage pas forcemment OK CBP7 il faut tester si bien ThEdit (forum)
  // BDU - 11/04/07 - FQ : 13921. Pas de saisie directe
  // BDU - 30/11/06, Ajout du type "Contact d'un autre tiers"
  // BDU - 15/01/2007. 13599 - Fait pointer les évènements sur des méthodes de ce source (plus de script)
  // BDU - 11/04/07 - FQ : 13921. Contrôle de la saisie manuelle

  LeContact.OnElipsisClick   := LeContactOnElipsisClick;
  LeContact.OnKeyPress       := ContactTouche;

  Tscontact.OnElipsisClick   := TousContactOnElipsisClick;
  TsContact.OnKeyPress       := ContactTouche;

  LeTiers.OnElipsisClick     := ClientClick;
  LeTiers.OnExit             := TiersExit;

  LaRessource.OnElipsisClick := RessourceClick;
  LaRessource.OnExit         := RessourceExit;

  TypeInterv.OnClick         := TypeIntervenantOnClick;

  {$IFDEF GIGI}
  // Pas de prospect si la GRC n'est pas sérialisée
  if not VH_GC.GRCSeria then
  begin
    if Assigned(TypeInterv) then
    begin
      X := TypeInterv.Values.IndexOf('PRO');
      TypeInterv.Items.Delete(X);
      TypeInterv.Values.Delete(X);
    end;
  end;

  // BDU - 11/04/07 - FQ : 13921
  // Pas de fournisseur si achat non sérialisé
  if not VH_GC.GAAchatSeria then
  begin
    if Assigned(TypeInterv) then
    begin
      X := TypeInterv.Values.IndexOf('FOU');
      TypeInterv.Items.Delete(X);
      TypeInterv.Values.Delete(X);
    end;
  end;
  {$ENDIF}

  if VH_GC.GCIfDefCEGID then SetControlText ('TAFT_LIENAFFTIERS', 'Statut');

end;

Procedure TOM_AFFTIERS.GetScreenObject;
begin

  TypeInterv    := THDBValcomboBox(GetControl('AFT_TYPEINTERV'));
  LeContact     := THEdit(GetControl('LECONTACT'));
  LaRessource   := THEdit(GetControl('LARESSOURCE'));
  LeTiers       := THEdit(GetControl('LETIERS'));
  TsContact     := THEdit(GetControl('TOUSCONTACT'));

  LeClient      := THEdit(GetControl('LECLIENT'));

  LibAffaire    := THLabel(GetControl('LIBAFFAIRE'));
  LibClient     := THLabel(GetControl('LIBCLIENT'));

end;

procedure TOM_AFFTIERS.Controlechamp(Champ, Valeur: String);
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  //
  if Champ='AFFAIRE'       then fStAffaire       := Valeur;
  if champ='LIBELLE'       then fStLibOrig       := Valeur;
  if Champ='TYPEINTERV'    then fStTypeInterv    := Valeur;
  if Champ='AFFAIREECOLE'  then fStAffaireEcole  := Valeur;
  if Champ='RANG'          then fStRang          := Valeur;

end;

{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.OnLoadRecord;
var StType : String;
begin
  inherited;

  //
  TypeInterv.Enabled := (ds.State in [dsinsert]);
  //
  LeContact.Text := '';
  //
  StType := TypeInterv.Value;

  //Affichage de la zone tiers...
  LeTiers.Visible     := (StType = 'CLI') or (StType = 'PRO') or (StType = 'FOU');
  LeTiers.Enabled     := (StType = 'CLI') or (StType = 'PRO') or (StType = 'FOU');
  Letiers.text        := GetField('AFT_TIERS');

  //Affichage de la zone ressource
  LaRessource.Visible := (StType = 'RES');
  LaRessource.Enabled := (StType = 'RES');
  LaRessource.text    := GetField('AFT_RESSOURCE');

  //Affichage de la zone contact (only one)
  LeContact.Visible := (StType = 'CON');
  LeContact.Enabled := (StType = 'CON');
  Lecontact.Text    := GetField('AFT_LIBELLE');

  //Affichage de la zone contact (All)
  TsContact.Visible := (StType = 'ACO');
  TsContact.Enabled := (StType = 'ACO');
  TsContact.text    := GetField('AFT_LIBELLE');
  //

end;
{$ENDIF !EAGLSERVER}

procedure TOM_AFFTIERS.OnNewRecord;
begin
  inherited;
  //Remise à zéro des zones écran....
  TypeInterv.Value  := '';
  LaRessource.Text  := '';
  LeTiers.text      := '';
  LeContact.Text    := '';
  TsContact.text    := '';
  //
  Setfield('AFT_LIBELLE', '');
  SetField ('AFT_TYPEINTERV',   TypeInterv.value);
  //
  Setfield ('AFT_TYPEORIG',     fStTypeOrig);
  Setfield ('AFT_NUMORIG',      fInNumOrig);
  Setfield ('AFT_INTERVENTION', fStLibOrig);
  SetField ('AFT_AUXILIAIRE',   fStAuxiliaire);
  //
  SetControlEnabled ('AFT_NUMORIG', false);
  //
  fStType := '';
  //
  if VH_GC.GCIfDefCEGID then SetField ('AFT_LIENAFFTIERS', 'INS');
  //
end;

procedure TOM_AFFTIERS.OnUpdateRecord;
var
  vQR: TQuery;
  IMax: integer;
  StSql: string;
begin
  inherited;

  SetField('AFT_TYPEINTERV', TypeInterv.value);

  //Contrôle cohérence des données....
  {1'Type d''intervenant obligatoire',}
  if TypeInterv.Value = '' then //
  begin
    LastError := 1;
    TypeInterv.SetFocus;
    LastErrorMsg := TraduitGa (TexteMessage [LastError] );
    Exit;
  end
  {2'Code intervenant obligatoire',}
  else if (LaRessource.text = '') and (LeTiers.text = '') and (GetField ('AFT_NUMEROCONTACT') = 0) then
  begin
    LastError := 2;
    // BDU - 13/03/07 - FQ : 13797
    if (TypeInterv.Value = 'CLI') or (TypeInterv.Value = 'PRO') or (TypeInterv.Value = 'FOU') then
      LeTiers.SetFocus
    else if TypeInterv.Value = 'RES' then
      LaRessource.SetFocus
    else if TypeInterv.Value = 'ACO' then // BDU - 30/11/06, Ajout du type "Contact d'un autre tiers"
      TsContact.SetFocus
    else
      LeContact.SetFocus;
    //
    LastErrorMsg := TraduitGa (TexteMessage [LastError] );
    Exit;
  end;

  if getfield ('AFT_RANG') = 0 then
  begin
    StSql := 'SELECT MAX(AFT_RANG) FROM AFFTIERS WHERE AFT_AFFAIRE="' + GetField ('AFT_AFFAIRE') +
      '" AND AFT_TYPEORIG="' + fStTypeOrig + '" AND AFT_NUMORIG=' + IntToStr (GetField ('AFT_NUMORIG'));
    vQR := OpenSQL (StSql, true);
    if not vQR.EOF then
      imax := vQR.Fields [0] .AsInteger + 1
    else
      iMax := 1;
    Ferme (vQR);
    SetField ('AFT_RANG', IMax);
  end;

  // BDU - 15/01/2007. 13599 - Mise à jours des zones contenant le libellé du contact
  if TypeInterv.Value = 'ACO' then
    TsContact.text := GetField('AFT_LIBELLE')
  else if TypeInterv.Value = 'CON' then
    LeContact.text := GetField('AFT_LIBELLE');

end;

procedure TOM_AFFTIERS.LoadAffaire;
var vSql      : string;
    vQR       : TQuery;
    iPosition : integer;
begin

  iPosition := ChargeCleAffaire (nil, THEdit (GetControl ('AFT_AFFAIRE1')), THEdit (GetControl ('AFT_AFFAIRE2')),
    THEdit (GetControl ('AFT_AFFAIRE3')), THEDIT (GetControl ('AFT_AVENANT')), nil, taConsult, fStAffaire, false);

  if iPosition > 0 then LibAffaire.left := iPosition;

  vSql := 'SELECT AFF_AFFAIRE, AFF_LIBELLE, T_TIERS, T_LIBELLE, T_AUXILIAIRE';

  if (fStTypeOrig = 'TAC') and (fInNumOrig <> 0) then vSql := vSql + ',ATA_LIBELLETACHE1';

  vSql := vSql + ' FROM TIERS,AFFAIRE';

  if (fStTypeOrig = 'TAC') and (fInNumOrig <> 0) then
    vSql := vSql + ' LEFT JOIN TACHE ON ATA_AFFAIRE="' + fStAffaire + '" AND ATA_NUMEROTACHE=' + IntToStr (fInNumOrig);

  vSql := vSql + ' WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
  vSql := vSql + ' AND AFF_TIERS = T_TIERS';

  vQR := OpenSql (vSql, True);

  if not vQR.Eof then
  begin
    LibAffaire.Caption  := vQR.FindField ('AFF_LIBELLE').AsString;
    LeClient.Text       := vQR.FindField ('T_TIERS').AsString;
    LibClient.Caption   := vQR.FindField ('T_LIBELLE').AsString;
    fStAuxiliaire       := vQR.FindField ('T_AUXILIAIRE').AsString;
    if (fStTypeOrig = 'TAC') and (fInNumOrig <> 0) then
      fStLibOrig        := vQR.FindField ('ATA_LIBELLETACHE1').AsString;
  end;

  ferme (vQR);

end;

{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.LeContactOnElipsisClick (Sender: TObject);
var StNumeroContact : string;
    StRang          : string;
    StSql           : string;
    iNumeroContact  : integer;
    Q               : TQuery;
begin

  StRang := 'T;' + fStAuxiliaire + ';';

  if GetField ('AFT_NUMEROCONTACT') <> 0 then StRang := StRang + IntToStr (GetField ('AFT_NUMEROCONTACT')) + ';';

  StNumeroContact := AGLLanceFiche ('YY', 'YYCONTACT', StRang, '', 'FROMSAISIE;FROMFORMATION;');
  //
  StNumeroContact := ReadTokenSt(StNumeroContact);
  iNumeroContact  := valeurI (StNumeroContact);
  //
  if (iNumeroContact = 0) then Exit;

  StSql := 'SELECT AFT_RANG FROM AFFTIERS WHERE AFT_AFFAIRE="' + GetField ('AFT_AFFAIRE') + '"' +
           ' AND AFT_TYPEORIG="' + fStTypeOrig + '" AND AFT_NUMORIG=' + IntToStr (fInNumOrig) +
           ' AND AFT_TYPEINTERV="CON" AND AFT_NUMEROCONTACT=' + IntToStr (iNumeroContact);

  {5'Ce contact est déjà affecté, voulez-vous l''enregistrer ?' }
  if (DS.State in [dsInsert] ) and Existesql(StSql) and (PGIAskAF (TexteMessage [5] , Ecran.caption ) <> mrYes) then Exit;

  if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;

  SetField ('AFT_NUMEROCONTACT', iNumeroContact);

  // BDU - 30/11/06, Ajout du type "Contact d'un autre tiers". Renseignement systématique du code tiers
  SetField('AFT_TIERS', TiersAuxiliaire(fStAuxiliaire, True));
  LeTiers.text := GetField('AFT_TIERS');

  // BDU - 15/01/2007. 13599 - Recherche les éléments du libellé
  StSql := 'SELECT C_NOM, C_PRENOM, C_CIVILITE FROM CONTACT ';
  StSql := StSQL + 'WHERE C_TYPECONTACT = "T" AND C_AUXILIAIRE ="' + fStAuxiliaire + '" ';
  StSQL := StSql + '  AND C_NUMEROCONTACT = ' + IntToStr(iNumeroContact);

  // BDU - 10/04/07 - FQ : 13919. Remplacer StNumeroContact par iNumeroContact
  Q := OpenSQl(StSql, True);
  try
    if not Q.EOF then
    begin
      SetField('AFT_LIBELLE', Q.Fields[2].AsString + ' ' + Q.Fields[0].AsString + ' ' + Q.Fields[1].AsString);
      LeContact.text := Trim(GetField('AFT_LIBELLE'));
    end
    else LeContact.text := '';
  finally
   Ferme(Q);
  end;

end;
{$ENDIF EAGLSERVER}

procedure TOM_AFFTIERS.TypeIntervenantOnClick (Sender: TObject);
begin

  if fStType <> TypeInterv.value then
  begin
    fStType := TypeInterv.value;
    // BDU - 13/03/07 - FQ : 13797
    LeTiers.Visible      := (fStType = 'CLI') or (fStType = 'PRO') or (fStType = 'FOU');
    LeTiers.Enabled      := LeTiers.Visible;
    //
    LaRessource.Visible  := (fStType = 'RES');
    LaRessource.Enabled  := LaRessource.Visible;
    //
    LeContact.Visible    := (fStType = 'CON');
    LeContact.Enabled    := LeContact.Visible;
    //
    TsContact.Visible    := (fStType = 'ACO');
    TsContact.enabled    := TsContact.Visible;
    //
    Setfield ('AFT_TIERS', '');
    Setfield ('AFT_RESSOURCE', '');
    Setfield ('AFT_LIBELLE', '');
    SetField ('AFT_NUMEROCONTACT', 0);
    //
    LeContact.text := '';
    TsContact.Text := '';
  end;

end;

procedure TOM_AFFTIERS.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_AFFTIERS.OnClose;
begin
  inherited;
end;

procedure TOM_AFFTIERS.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_AFFTIERS.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_AFFTIERS.OnAfterDeleteRecord;
begin
  inherited;
end;

procedure TOM_AFFTIERS.ClientClick(Sender: TObject);
Var S     : String;
    Q     : TQuery;
begin

  {$IFNDEF ERADIO}
    // BDU - 13/03/07 - FQ : 13797
    S := TypeInterv.Value;
    S := DispatchRecherche(Letiers, 2, 'T_NATUREAUXI="' + S + '"', '', '');
  {$ELSE  !ERADIO}
    S := '';
  {$ENDIF !ERADIO}

  if S <> '' then
  begin
    // BDU - 15/01/2007. 13599 - Récupération des éléments du libellé
    Q := OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS = "' + S + '"', True);
    try
      if not Q.EOF then
      Begin
        SetField('AFT_LIBELLE', Q.Fields[0].AsString);
        SetField('AFT_TIERS',   S);
        LeTiers.Text := S;
      end
      else
      Begin
        SetField('AFT_LIBELLE', '');
        SetField('AFT_TIERS',   '');
        LeTiers.Text := '';
      end;
    finally
      Ferme(Q);
    end;
  end;

end;

procedure TOM_AFFTIERS.RessourceClick(Sender: TObject);
var
  Q: TQuery;
  S: String;
begin

  {$IFNDEF ERADIO}
    S := DispatchRecherche(LaRessource, 3, '', '', '');
  {$ELSE  !ERADIO}
    S := '';
  {$ENDIF !ERADIO}

  if S <> '' then
  begin
    // BDU - 15/01/2007. 13599 - Récupération des éléments du libellé
    Q := OpenSQL('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE = "' + S + '"', True);
    try
      if not Q.EOF then
      Begin
        SetField('AFT_LIBELLE', Q.Fields[0].AsString);
        SetField('AFT_RESSOURCE', S);
        LaRessource.Text := S;
      end
      else
      Begin
        SetField('AFT_LIBELLE',   '');
        SetField('AFT_RESSOURCE', '');
        LaRessource.Text := '';
      end;
    finally
      Ferme(Q);
    end;
  end;

end;

// BDU - 30/11/06, Ajout du type "Contact d'un autre tiers"
{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.TousContactOnElipsisClick (Sender: TObject);
var
  Retour,
  Auxiliaire: string;
  NumeroContact: integer;
  Q: TQuery;
begin

  Retour := AGLLanceFiche ('AFF', 'AFCONTACT_MUL', '', '', 'CONSULTATION;CONTACT_AFFAIRE');

  try
    // La première valeur n'est pas utilisée
    ReadTokenSt(Retour);
    // L'auxiliaire est en deuxième position
    Auxiliaire := ReadTokenSt(Retour);
    // Le numéro de contact est en troisième position
    NumeroContact := StrToInt(ReadTokenSt(Retour));
  except
    NumeroContact := 0;
  end;

  if (NumeroContact <> 0) then
  begin
    if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
    // Attention, l'affectation de AFT_TIERS doit être avant celle de AFT_NUMEROCONTACT
    SetField('AFT_TIERS', TiersAuxiliaire(Auxiliaire, True));
    SetField ('AFT_NUMEROCONTACT', NumeroContact);
    //
    LeTiers.Text := GetField('AFT_TIERS');
    // BDU - 15/01/2007. 13599 - Récupération des éléments du libellé
    // BDU - 01/10/07 - FQ : 13963. Ajout du code du tiers
    Q := OpenSQl('SELECT C_NOM, C_PRENOM, C_CIVILITE, C_TIERS FROM CONTACT ' +
      'WHERE C_TYPECONTACT = "T" AND C_TIERS = "' + GetField('AFT_TIERS') +
      '" AND C_NUMEROCONTACT = ' + IntToStr(NumeroContact), True);
    try
      if not Q.EOF then
      begin
        // BDU - 01/10/07 - FQ : 13963. Ajout du code du tiers
        SetField('AFT_LIBELLE', Q.Fields[3].AsString + ' - ' + Q.Fields[2].AsString + ' ' + Q.Fields[0].AsString + ' ' + Q.Fields[1].AsString);
        TsContact.text := GetField('AFT_LIBELLE');
      end
      else
        TsContact.text := '';
    finally
      Ferme(Q);
    end;
  end;

end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.ContactTouche(Sender: TObject; var Key: char);
begin
  Key := #0;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.RessourceExit(Sender: TObject);
var
  S: String;
  Q: TQuery;
begin

  if DS.State in [dsEdit, dsInsert] then
  begin
    S := UpperCase(Trim(LaRessource.text));

    if S = '' then
    begin
      SetControlText('AFT_LIBELLE', '');
      Exit;
    end;

    S := 'SELECT ARS_LIBELLE, ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE = "' + S + '"';
    Q := OpenSQL(S, True);
    try
      if not Q.EOF then
      begin
        SetField('AFT_LIBELLE', Q.Fields[0].AsString);
        SetField('AFT_RESSOURCE', Q.Fields[1].AsString);
        LaRessource.Text := Getfield('AFT_RESSOURCE');
      end
      else
      begin
        SetField('AFT_LIBELLE', '');
        SetField('AFT_RESSOURCE', '');
        LaRessource.Text := Getfield('AFT_RESSOURCE');
        LaRessourcE.SetFocus;
      end;
    finally
      Ferme(Q);
    end;
  end;

end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure TOM_AFFTIERS.TiersExit(Sender: TObject);
var
  S: String;
  Q: TQuery;
begin

  if DS.State in [dsEdit, dsInsert] then
  begin
    S := UpperCase(Trim(LeTiers.Text));

    if S = '' then
    begin
      SetControlText('AFT_LIBELLE', '');
      Exit;
    end;

    S := 'SELECT T_LIBELLE, T_TIERS FROM TIERS WHERE T_TIERS = "' + S + '"';
    S := S + ' AND T_NATUREAUXI = "' + THDBValcomboBox (GetControl ('AFT_TYPEINTERV')).Value + '"';
    Q := OpenSQL(S, True);
    try
      if not Q.EOF then
      begin
        SetField('AFT_LIBELLE', Q.Fields[0].AsString);
        SetField('AFT_TIERS', Q.Fields[1].AsString);
        LeTiers.text := GetField('AFT_TIERS');
      end
      else
      begin
        SetField('AFT_LIBELLE', '');
        SetField('AFT_TIERS', '');
        LeTiers.text := '';
        LeTiers.SetFocus;
      end;
    finally
      Ferme(Q);
    end;
  end;

end;
{$ENDIF EAGLSERVER}

initialization
  registerclasses ([TOM_AFFTIERS] );
end.
