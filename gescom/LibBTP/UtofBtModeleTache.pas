{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : Saisie des Modèles de tâches
Mots clefs ... : TABLE : AFMODELETACHE - Fiche BTTACHEMODELE
******************************************************************}
unit UtofBtModeleTache;

interface
uses StdCtrls,
  UTOF,
  Controls,
  Classes,
  HTB97,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}
  dbtables {BDE},
{$ELSE}
  uDbxDataSet,
{$ENDIF}
  Fiche,
  FE_Main,
{$ELSE}
  eFiche,
  MaineAGL,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  UtilConfid,
  UtilGC,
  AfUtilArticle,
  UtilTaches,
  UtilRessource,
  paramsoc,
  Vierge,
  AGLInitGC,
  Dicobtp;

type
  TOF_BTMODELETACHE = class (TOF)
    procedure OnNew                   ; override;
    procedure OnDelete                ; override;
    procedure OnUpdate                ; override;
    procedure OnLoad                  ; override;
    procedure OnArgument(StArgument: string); override;
    procedure OnClose                 ; override;

  private

    fAction             : TActionFiche;
    fStatut             : TActionFiche; // statut interne de la fiche (pour enregistrement)

    fTobModele          : Tob;
    TobTacheJour        : TOB;

    fBoModifRegles      : Boolean;

    CodeModele          : String;
    fStLibelleParDefaut : string;

    BT                 : TToolBarButton97;

    RB_MODEGENE1       : TRadioButton;
    RB_MODEGENE2       : TRadioButton;
    RB_QUOTIDIENNE     : TRadioButton;
    RB_NBINTERVENTION  : TRadioButton;
    RB_HEBDOMADAIRE    : TRadioButton;
    RB_ANNUELLE        : TRadioButton;
    RB_MENSUELLE       : TRadioButton;
    RB_MOISMETHODEJOUR : TRadioButton;
    RB_MOISMETHODELE   : TRadioButton;
    RB_MOISMETHODETOUTLES : TRadioButton;

    PC_FREQUENCE       : TPageControl;

    TS_QUOTIDIENNE     : TTAbSheet;
    TS_NBINTERVENTION  : TTAbSheet;
    TS_HEBDOMADAIRE    : TTAbSheet;
    TS_ANNUELLE        : TTAbSheet;
    TS_MENSUELLE       : TTAbSheet;

    LibUnitTemps       : THLabel;
    UniteTemps         : THValComboBox;

    procedure AFM_UNITETEMPSOnChange(Sender: TObject);
    procedure AlimTob (NoJour: integer);
    procedure ArticleOnElipsisClick(sender: Tobject);
    procedure ArticleOnexit(sender: Tobject);
    procedure BDupliqueOnClick(Sendre: TObject);
    procedure BInsertOnClick(Sendre: TObject);

    procedure fRB_ANNUELLEOnClick(Sender: TObject);
    procedure fED_DATEPERIODDEBOnExit (SEnder: TObject);
    procedure fED_DATEPERIODFINOnExit (SEnder: TObject);
    procedure fRB_HEBDOMADAIREOnClick(Sender: TObject);
    procedure fRB_MENSUELLEOnClick(Sender: TObject);
    procedure fRB_MODEGENE1OnClick(Sender: TObject);
    procedure fRB_MODEGENE2OnClick(Sender: TObject);
    procedure fRB_MOISMETHODEJOUROnClick(SEnder: TObject);
    procedure fRB_MOISMETHODELEOnClick(SEnder: TObject);
    procedure fRB_MOISMETHODETOUTLESOnClick(SEnder: TObject);
    procedure fRB_NBINTERVENTIONOnClick(Sender: TObject);
    procedure fRB_QUOTIDIENNEOnClick(Sender: TObject);
    procedure MODELETACHEOnexit(Sender: TObject);
    procedure TypeActionOnExit(Sender: TObject);

    procedure ControleChamp(Champ, Valeur: String);
    procedure GestionEcran;
    procedure InitNewModele;
    procedure InitNewRegles;
    procedure LectureTypeAction(TypeAction: string);
    procedure RaffraichirArticle;
    procedure ModifZoneExit(Sender: TObject);
    procedure GetRegles;
    procedure bDeleteOnClick(SEnder: TObject);

  end;

implementation

uses TntStdCtrls;
const
  TexteMessage: array [1..10] of string = (
    {1}  'La date de début est supérieure à la date de fin.'
    {2}, 'L''article est obligatoire.'
    {3}, 'L''article n''existe pas.'
    {4}, 'La durée d''une intervention doit être positive.'
    {5}, 'La famille de tâche est obligatoire.'
    {6}, 'Ce modèle de tâche a été utilisé en création de tâche,#13#10 Confirmez-vous sa suppression ?'
    {7}, 'La saisie du champ suivant est obligatoire : '
    {8}, 'Le code modèle de tâche est obligatoire '
    {9}, 'Confirmez vous la suppression de la tâche ?'
    {10}, 'Vous ne pouvez pas utiliser ce type d''action (réservé)'
    );

{ TOF_BTMODELETACHE }

procedure TOF_BTMODELETACHE.OnArgument(StArgument: string);
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  inherited;

  fAction := TaCreat;
  fStatut := TaCreat;

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(stArgument)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
        begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
          begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
          ControleChamp(champ, valeur);
	        end
        end;
      Critere := (Trim(ReadTokenSt(stArgument)));
    end;

  //Chargement du code affaire
  if GetControlText('AFM_MODELETACHE') <> '' then CodeModele :=GetControlText('AFM_MODELETACHE');

  GestionEcran;
  //
  RB_MOISMETHODEJOUR.OnClick := fRB_MOISMETHODEJOUROnClick;
  RB_MOISMETHODELE.OnClick := fRB_MOISMETHODELEOnClick;
  RB_MOISMETHODETOUTLES.OnClick := fRB_MOISMETHODETOUTLESOnClick;
  //

end;

Procedure TOF_BTMODELETACHE.ControleChamp(Champ : String;Valeur : String);
Begin

  if Champ = 'ACTION' then
	   Begin
     if Valeur = 'CREATION' Then
        fAction := TaCreat
     else if Valeur = 'MODIFICATION' then
        fAction := TaModif
     else
	      fAction := TaConsult
  end;

  if Champ ='CODEMODELE' then CodeModele :=valeur;

end;


procedure TOF_BTMODELETACHE.OnDelete;
begin
  inherited;
end;

procedure TOF_BTMODELETACHE.bDeleteOnClick(SEnder: TObject);
var Action : TCloseAction;
begin

  //Confirmez vous la suppression de la tâche ?
  if (PGIAskAF( TexteMessage[9], ecran.Caption) = mrno) then exit;

  //6, 'Ce modèle de tâche a été utilisé en création de tâche,#13#10 Confirmez-vous sa suppression ?'}
  if ExisteSQL('SELECT ATA_AFFAIRE FROM TACHE WHERE ATA_MODELETACHE="' + codemodele + '"') and
    (PGIAskAF (TexteMessage [6] , Ecran.Caption) <> mrYes) then
    begin
    LastError := 6;
    exit;
    end;

  // il faut aussi détruire la table tachejour et la table mémo associée
  ExecuteSql ('DELETE FROM AFMODELETACHE WHERE AFM_MODELETACHE="' + CodeModele + '"');
  ExecuteSql ('DELETE FROM AFTACHEJOUR WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE="' + CodeModele + '"');
  ExecuteSql ('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="AFM" AND LO_IDENTIFIANT="' + CodeModele + '"');

  Ecran.Close;

end;

procedure TOF_BTMODELETACHE.OnLoad;
Var Tobjour : Tob;
    TobDet  : Tob;
    StrSQL  : String;
    QQ      : TQuery;
    ii      : integer;
    annee   : Word;
    jj      : Word;
    mm      : Word;
    Qte     : Double;
begin
  inherited;

  SetActiveTabSheet('PGENERAL');

  AppliquerConfidentialite(Ecran, '');

  InitNewRegles;

  if fAction = taCreat then
     Begin
     OnNew;
     Exit;
     end;

  //si modification chargment de la tob Modele de Tache
  StrSQL := 'SELECT * FROM AFMODELETACHE WHERE AFM_MODELETACHE ="' + CodeModele + '"';
  QQ := OpenSql (StrSQL,true);

  if QQ.eof then
     begin
     ferme(QQ);
     fAction := TaCreat;
     OnNew;
     Exit;
     end;

  fTobModele.selectDB('',QQ);
  ferme (QQ);

  //Chargement de la Tob des Jours
  TobJour := Tob.create ('_AFTACHEJOUR_', nil, -1);
  StrSQL := 'SELECT * FROM AFTACHEJOUR WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE="' + CodeModele + '"';
  TobJour.LoadDetailDBFromSQL ('_AFTACHEJOUR_', StrSQL, True);

  //chargement de l'écran PGENERAL
  SetControlText('AFM_MODELETACHE', CodeModele);
  SetcontrolProperty('AFM_MODELETACHE','Enabled', false);

  SetControlText('AFM_FAMILLETACHE',  Ftobmodele.GetString('AFM_FAMILLETACHE'));
  SetControlText('AFM_TYPEARTICLE',   Ftobmodele.GetString('AFM_TYPEARTICLE'));
  SetControlText('AFM_CODEARTICLE',   Ftobmodele.GetString('AFM_CODEARTICLE'));
  SetControlText('AFM_LIBELLETACHE1', Ftobmodele.GetString('AFM_LIBELLETACHE1'));
  SetControlText('AFM_LIBELLETACHE2', Ftobmodele.GetString('AFM_LIBELLETACHE2'));

  SetControlText('AFM_BTETAT',        Ftobmodele.GetString('AFM_BTETAT'));
  LectureTypeAction(GetControlText('AFM_BTETAT'));

  SetControlText('AFM_DATEDEBPERIOD', Ftobmodele.GetString('AFM_DATEDEBPERIOD'));
  SetControlText('AFM_DATEFINPERIOD', Ftobmodele.GetString('AFM_DATEFINPERIOD'));

  UniteTemps.Value := Ftobmodele.GetString('AFM_UNITETEMPS');
  LibUnitTemps.Caption := UniteTemps.Text;

  SetControlText('AFM_DESCRIPTIF',    Ftobmodele.GetString('AFM_DESCRIPTIF'));

  //Chargement de l'écran des règles
  Qte := strtofloat(Ftobmodele.GetValue('AFM_QTEINTERVENT'));

  SetControlText('AFM_QTEINTERVENT',  FloatToStr(Qte));
  SetControlText('AFM_DATEANNUELLE',  Ftobmodele.GetString('AFM_DATEANNUELLE'));
  SetControlText('AFM_METHODEDECAL',  Ftobmodele.GetString('AFM_METHODEDECAL'));
  SetControlText('AFM_NBJOURSDECAL',  Ftobmodele.GetString('AFM_NBJOURSDECAL'));

  if FtobModele.GetValue('AFM_MODEGENE') = 1 then
     RB_MODEGENE1.Checked := True // au plus tôt
  else
     RB_MODEGENE2.Checked := True; // au plus tard

  if FtobModele.GetValue('AFM_PERIODICITE') = 'Q'   then
  begin
    RB_QUOTIDIENNE.Checked := true;
    SetControltext('AFM_JOURINTERVAL', FtobModele.GetValue('AFM_JOURINTERVAL '));
  end;

  if FtobModele.GetValue('AFM_PERIODICITE') = 'NBI' then
  begin
    SetControlText('AFM_NBINTERVENTION',Ftobmodele.GetString('AFM_NBINTERVENTION'));
    RB_NBINTERVENTION.Checked := true;
  end;

  if FtobModele.GetValue('AFM_PERIODICITE') = 'S'   then
  Begin
    RB_HEBDOMADAIRE.Checked := true;
    SetControlText('AFM_SEMAINEINTERV',Ftobmodele.GetString('AFM_SEMAINEINTERV'));
    for ii := 0 to Tobjour.detail.count - 1 do
    begin
      TobDet := TobJour.detail [ii] ;
      TCheckBox(GetControl('AFM_JOUR' + IntToStr (TobDet.getvalue ('ATJ_JOURAPLANIF')) + 'H')).checked := True;
    end;
  end;

  if FtobModele.GetValue('AFM_PERIODICITE') = 'A' then
  begin
    RB_ANNUELLE.Checked := true;
    Decodedate (FtobModele.GetValue('AFM_DATEANNUELLE'), annee, mm, jj);
    SetControlText ('JOURAN', IntToStr (jj));
    SetControlText ('MOISAN', IntTostr (mm));
    SetControlText ('AFM_ANNEENB', FtobModele.GetValue('AFM_ANNEENB'));
  end;

  if FtobModele.GetValue('AFM_PERIODICITE') = 'M' then
  begin
    RB_MENSUELLE.Checked := true;
    if FtobModele.GetValue('AFM_MOISMETHODE') = '1' then
    begin
      RB_MOISMETHODEJOUR.Checked := True;
      SetControlText('AFM_MOISJOURFIXE', FtobModele.GetValue('AFM_MOISJOURFIXE'));
      SetControlText('AFM_MOISFIXE0', FtobModele.GetValue('AFM_MOISFIXE'));
    end
    else if FtobModele.GetValue('AFM_MOISMETHODE') = '2' then
    begin
      RB_MOISMETHODELE.Checked := True;
      SetControlText('AFM_MOISFIXE1',    FtobModele.GetValue('AFM_MOISFIXE'));
      THValComboBox(GetControl('AFM_MOISSEMAINE1')).Value := FtobModele.GetValue('AFM_MOISSEMAINE');
      THValComboBox(GetControl('AFM_MOISJOURLIB')).Value := FtobModele.GetValue('AFM_MOISJOURLIB');
    end
    else if FtobModele.GetValue('AFM_MOISMETHODE') = '3' then
    begin
      RB_MOISMETHODETOUTLES.Checked := True;
      SetControlText('AFM_MOISFIXE2', FtobModele.GetValue('AFM_MOISFIXE'));
      THValComboBox(GetControl('AFM_MOISSEMAINE2')).Value := FtobModele.GetValue('AFM_MOISSEMAINE');
      for ii := 0 to Tobjour.detail.count - 1 do
      begin
        TobDet := TobJour.detail [ii] ;
        TCheckBox(GetControl('AFM_JOUR' + IntToStr (TobDet.getvalue ('ATJ_JOURAPLANIF')) + 'M')).checked := True;
      end;
    end;
  end;

  Tobjour.free;

  // pour que le chargement des règles ne les mettent pas modifiées
  fBoModifRegles := false;
end;

procedure TOF_BTMODELETACHE.OnNew;
begin
  inherited;

  SetActiveTabSheet('PGeneral');

  // pour l'instant, on initialise avec Heure comme unité de saisie
  UniteTemps.Value := 'H';
  LibUnitTemps.Caption := UniteTemps.Text;
  
  SetControlText ('AFM_FAMILLETACHE', GetParamsocSecur('SO_BTFAMILLEDEF', ''));

  SetControlProperty('fRB_MODEGENE1', 'Checked', true);
  SetControlProperty('fRB_MENSUELLE', 'Checked', true);
  SetControlProperty('fRB_MOISMETHODEJOUR', 'Checked', true);

  SetControlText('AFM_MOISMETHODE', '1');
  SetControlText('AFM_MOISFIXE', '1');
  SetControlText('AFM_MOISJOURLIB', 'J1');
  SetControlText('AFM_MOISJOURFIXE', '1');

  SetControlText('AFM_DESCRIPTIF', '');

  fRB_MensuelleOnClick(self);

  SetFocusControl('AFM_MODELETACHE');

  InitNewRegles;

  InitNewModele;

end;


procedure TOF_BTMODELETACHE.OnUpdate;
Var vStCodeArticle  : String;
    vStArticle      : String;
    vStTypeArticle  : String;
    vStFacturable   : String;
    vStUnite        : String;
    vStLibelle      : string;
    vStStatutPla    : String;
    Tobdet          : tob;
    QQ              : Tquery;
    ii              : Integer;
    indice          : integer;
    NomChamp        : string;
    PR              : Double;
    PV              : double;
    Qte             : Integer;
begin
  inherited;

  vStCodeArticle := trim (GetControlText('AFM_CODEARTICLE'));
  vStTypeArticle := GetControlText('AFM_TYPEARTICLE');
  vStUnite := UniteTemps.value;
  Qte := StrToInt(GetControlText('AFM_QTEINTERVENT'));

	if GetControlText('AFM_BTETAT') = '' then
  begin
     PgiInfo('Veuillez définir le type d''action');
     THEdit(GetControl('AFM_BTETAT')).SetFocus;
  	 TFVierge(Ecran).ModalResult := 0;
     exit;
  end;

 //contrôle de la cohérence des zones
  if GetControlText('AFM_MODELETACHE') = '' then
     begin
     LastError := 8;
     LastErrorMsg := TraduitGa(TexteMessage [LastError] );
     SetFocusControl ('AFM_MODELETACHE');
  	 TFVierge(Ecran).ModalResult := 0;
     exit;
     end;

  //1'La date de début est supérieure à la date de fin.'
  if strToDate (getControlText('AFM_DATEDEBPERIOD')) > strToDate (getControlText('AFM_DATEFINPERIOD')) then
     begin
     LastError := 1;
     LastErrorMsg := TraduitGa (TexteMessage [LastError] );
     SetFocusControl ('AFM_DATEFINPERIOD');
  	 TFVierge(Ecran).ModalResult := 0;
     exit;
     end;

  //2, 'L'article est obligatoire.'
  {if vStCodeArticle = '' then
     begin
     LastError := 2;
     LastErrorMsg := TraduitGa (TexteMessage [LastError] );
     SetFocusControl ('AFM_CODEARTICLE');
     exit;
     end;}

  //3, 'L'article n'existe pas.'
  if vStCodeArticle <> '' then
     Begin
     if not controleCodeArticle (vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle, PR, PV, vStStatutPla) then
        begin
        LastError := 3;
        LastErrorMsg := TraduitGa (TexteMessage [LastError] );
        SetFocusControl ('AFM_CODEARTICLE');
  	 		TFVierge(Ecran).ModalResult := 0;
        exit;
        end;
     end;

  //4, 'La durée d'une intervention doit être positive.'
  if Qte <= 0 then
      begin
      LastError := 4;
      LastErrorMsg :=TraduitGa (TexteMessage [LastError] );
      SetFocusControl ('AFM_QTEINTERVENT');
  	 	TFVierge(Ecran).ModalResult := 0;
      exit;
      end;

  //7, 'La saisie du champ suivant est obligatoire : '}
  NomChamp := VerifierChampsObligatoires (Ecran, '');
  if NomChamp <> '' then
     begin
     NomChamp := ReadTokenSt (NomChamp);
     SetFocusControl (NomChamp);
     LastError := 7;
     LastErrorMsg := TexteMessage [LastError] + champToLibelle (NomChamp);
  	 TFVierge(Ecran).ModalResult := 0;
     end;

  TobTacheJour := Tob.create ('AFTACHEJOUR', nil, -1);

  SetActiveTabSheet('PGENERAL');

  //chargement de la tob modèle de tâche
  FtobModele.PutValue('AFM_MODELETACHE',    GetControlText('AFM_MODELETACHE'));
  FtobModele.PutValue('AFM_FAMILLETACHE',   GetControlText('AFM_FAMILLETACHE'));
  FtobModele.PutValue('AFM_TYPEARTICLE',    vStTypeArticle);
  FtobModele.PutValue('AFM_ARTICLE',        vStArticle);
  FtobModele.PutValue('AFM_CODEARTICLE',    vStCodeArticle);
  Ftobmodele.PutValue('AFM_BTETAT',         GetControlText('AFM_BTETAT'));
  FtobModele.PutValue('AFM_DATEDEBPERIOD',  StrToDate (GetControlText ('AFM_DATEDEBPERIOD')));
  FtobModele.PutValue('AFM_DATEFINPERIOD',  StrToDate (GetControlText ('AFM_DATEFINPERIOD')));
  FtobModele.PutValue('AFM_UNITETEMPS',     UniteTemps.Value);
  FtobModele.PutValue('AFM_QTEINTERVENT',   Qte);
  FtobModele.PutValue('AFM_METHODEDECAL',   GetControlText ('AFM_METHODEDECAL'));
  FtobModele.PutValue('AFM_NBJOURSDECAL',   GetControlText ('AFM_NBJOURSDECAL'));

  FtobModele.PutValue('AFM_LIBELLETACHE1',  GetControlText ('AFM_LIBELLETACHE1'));
  FtobModele.PutValue('AFM_LIBELLETACHE2',  GetControlText ('AFM_LIBELLETACHE2'));
  FtobModele.PutValue('AFM_DESCRIPTIF',     GetControlText ('AFM_DESCRIPTIF'));

  GetRegles;

  FTobModele.InsertOrUpdateDb(False);

  ExecuteSql ('DELETE FROM AFTACHEJOUR WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE="' + CodeModele + '"');
  indice := 0;
  QQ := OpenSQL ('SELECT MAX(ATJ_JOURNUMERO) FROM AFTACHEJOUR', true);
  if not QQ.Eof then Indice := QQ.Fields [0].AsInteger + 1;
  Ferme (QQ);
  for ii := 0 to TobTacheJour.detail.count - 1 do
      begin
      Tobdet := TobTacheJour.detail [ii] ;
      Tobdet.putvalue ('ATJ_JOURNUMERO', Indice);
      Inc (indice);
      end;

  TobTacheJour.InsertOrUpdateDb (False);
  TobTacheJour.free;

  fStatut := taConsult;

end;

procedure TOF_BTMODELETACHE.GetRegles;
var Days : Word;
    Mnth : Word;
    Annee: Word;
    Mois : String;
    Jour : String;
begin

  if RB_MODEGENE1.Checked then
     FtobModele.PutValue('AFM_MODEGENE', 1)
  else
     FtobModele.PutValue('AFM_MODEGENE', 2);

  FtobModele.putvalue('AFM_QTEINTERVENT', GetControlText('AFM_QTEINTERVENT'));

  if RB_QUOTIDIENNE.Checked then
  begin
     FtobModele.PutValue('AFM_PERIODICITE', 'Q');
     FtobModele.PutValue('AFM_JOURINTERVAL',  GetControlText('AFM_JOURINTERVAL'));
  end
  else if RB_NBINTERVENTION.Checked then
  begin
     FtobModele.PutValue('AFM_PERIODICITE', 'NBI');
     FtobModele.PutValue('AFM_NBINTERVENTION', GetControlText ('AFM_NBINTERVENTION'));
  end
  else if RB_HEBDOMADAIRE.Checked then
     begin
     FtobModele.PutValue('AFM_PERIODICITE', 'S');
     FtobModele.PutValue('AFM_SEMAINEINTERV',  GetControlText('AFM_SEMAINEINTERV'));
     if TCheckBox(GetControl('AFM_JOUR1H')).Checked then AlimTob(1);
     if TCheckBox(GetControl('AFM_JOUR2H')).Checked then AlimTob(2);
     if TCheckBox(GetControl('AFM_JOUR3H')).Checked then AlimTob(3);
     if TCheckBox(GetControl('AFM_JOUR4H')).Checked then AlimTob(4);
     if TCheckBox(GetControl('AFM_JOUR5H')).Checked then AlimTob(5);
     if TCheckBox(GetControl('AFM_JOUR6H')).Checked then AlimTob(6);
     if TCheckBox(GetControl('AFM_JOUR7H')).Checked then AlimTob(7);
     end
  else if RB_ANNUELLE.Checked then
     begin
     DecodeDate(now, Annee, Mnth, Days);
     //Annee := Years(now);
     Mois  := GetControlText('MOISAN');
     Jour  := GetControlText('JOURAN');
     Mnth  := StrToInt(Mois);
     Days  := StrToInt(Jour);
     FtobModele.PutValue('AFM_PERIODICITE', 'A');
     FtobModele.PutValue('AFM_ANNEENB', GetControlText('AFM_ANNEENB'));
     //mis d'une année fictive dans la date .. seul jour et mois utilisé
     FtobModele.PutValue('AFM_DATEANNUELLE', Encodedate(Annee, Mnth, days));
     end
  else if RB_MENSUELLE.Checked then
     begin
     FtobModele.PutValue('AFM_PERIODICITE', 'M');
     if RB_MOISMETHODEJOUR.Checked then
        begin
        FtobModele.PutValue('AFM_MOISJOURFIXE', GetControlText('AFM_MOISJOURFIXE'));
        FtobModele.PutValue('AFM_MOISMETHODE', '1');
        FtobModele.PutValue('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE0'));
        end
     else if RB_MOISMETHODELE.Checked then
        begin
        FtobModele.PutValue('AFM_MOISMETHODE', '2');
        FtobModele.PutValue('AFM_MOISSEMAINE', THValComboBox(GetControl('AFM_MOISSEMAINE1')).value);
        FtobModele.PutValue('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE1'));
        FtobModele.PutValue('AFM_MOISJOURLIB', THValComboBox(GetControl('AFM_MOISJOURLIB')).Value);
        end
     else if RB_MOISMETHODETOUTLES.Checked then
        begin
        FtobModele.PutValue('AFM_MOISMETHODE', '3');
        FtobModele.PutValue('AFM_MOISSEMAINE', THValComboBox(GetControl('AFM_MOISSEMAINE2')).value);
        FtobModele.PutValue('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE2'));
        if TCheckBox(GetControl('AFM_JOUR1M')).Checked then AlimTob(1);
        if TCheckBox(GetControl('AFM_JOUR2M')).Checked then AlimTob(2);
        if TCheckBox(GetControl('AFM_JOUR3M')).Checked then AlimTob(3);
        if TCheckBox(GetControl('AFM_JOUR4M')).Checked then AlimTob(4);
        if TCheckBox(GetControl('AFM_JOUR5M')).Checked then AlimTob(5);
        if TCheckBox(GetControl('AFM_JOUR6M')).Checked then AlimTob(6);
        if TCheckBox(GetControl('AFM_JOUR7M')).Checked then AlimTob(7);
        end;
     end;

  FtobModele.PutValue('RAPPTPRCAL', Valeur(GetControlText('RAFPTPRCALC')));
  FtobModele.PutValue('RAPPTPVCAL', Valeur(GetControlText('RAFPTVENTEHTCALC')));

end;

procedure TOF_BTMODELETACHE.AlimTob(NoJour: integer);
var
  tobdet: tob;
begin

  Tobdet := Tob.Create ('AFTACHEJOUR', TobTacheJour, -1);

  TOBDet.PutValue ('ATJ_Affaire', '');
  TOBDet.PutValue ('ATJ_NumeroTache', 0);
  TOBDet.PutValue ('ATJ_JourHeureDeb', 0);
  TOBDet.PutValue ('ATJ_JourNbHeures', 0);
  TOBdet.PutValue ('ATJ_JourAplanif', noJour);
  TOBdet.PutValue ('ATJ_ModeleTache', Getcontroltext('AFM_ModeleTache'));
  TOBdet.PutValue ('ATJ_TypeJour', 'MOD');

end;

Procedure TOF_BTMODELETACHE.GestionEcran;
Var Bt    : TToolBarButton97;
    Ed    : THEdit;
    VArtic: THEdit;
Begin

  //création de la Tob Modele de taches
  fTobModele := Tob.create('AFMODELETACHE', nil, -1);

  // Gestion des Boutons
  Bt := TToolBarButton97(GetControl('BInsert')); //création d'un nouvel Enregistrement
  if (Bt <> nil) then Bt.Onclick := BInsertOnClick;

  Bt := TToolBarButton97(GetControl('BDuplique')); // Duplication du modéle de tache sélectionné
  if (Bt <> nil) then
    Bt.Onclick := bDupliqueOnClick;


  LibUnitTemps := ThLabel(GetControl('ATA_UNITETEMPS2')); //Libellé de l'unité de temps

  //Gestion des Liste de choix...
  vArtic := THedit(getcontrol('AFM_CODEARTICLE'));
  if vArtic is THEdit then
     begin
     vArtic.OnExit := ArticleOnexit;
     vArtic.OnElipsisClick := ArticleOnElipsisClick;
     end;

  Bt := TToolBarButton97(GetControl('BDelete'));
  if (Bt <> nil) then
    Bt.Onclick := BDeleteOnClick;

  Ed := THedit(GetControl('AFM_MODELETACHE'));
  Ed.OnExit := MODELETACHEOnexit;

  Ed := THedit(GetControl('AFM_BTETAT'));
  Ed.OnExit := TypeActionOnexit;

  Ed := THedit(GetControl('AFM_DATEDEBPERIOD'));
  Ed.OnExit := fED_DATEPERIODDEBOnExit;

  Ed := THedit(GetControl('AFM_DATEFINPERIOD'));
  Ed.OnExit := fED_DATEPERIODFINOnExit;

  UniteTemps := THValComboBox(GetControl('AFM_UNITETEMPS'));
  if UniteTemps <> nil then
  begin
    UniteTemps.OnChange := AFM_UNITETEMPSOnChange;
  end;

  //gestion de la modification de la Regle de calcul
  Ed := THEdit(GetControl('AFM_JOURINTERVAL'));
  Ed.OnExit := ModifZoneExit;

  //Semaine
  Ed := THEdit(GetControl('AFM_SEMAINEINTERV'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR1H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR2H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR3H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR4H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR5H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR6H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR7H'));
  Ed.OnExit := ModifZoneExit;

  //Mois
  Ed := THEdit(GetControl('AFM_MOISJOURFIXE'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISFIXE0'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISFIXE1'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISFIXE2'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISSEMAINE1'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISSEMAINE2'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_MOISJOURLIB'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR1M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR2M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR3M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR4M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR5M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR6M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_JOUR7M'));
  Ed.OnExit := ModifZoneExit;

  //Année
  Ed := THEdit(GetControl('JOURAN'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('MOISAN'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('AFM_ANNEENB'));
  Ed.OnExit := ModifZoneExit;

  //Nombre d'intervention
  Ed := THEdit(GetControl('AFM_NBINTERVENTION'));
  Ed.OnExit := ModifZoneExit;

  // gestion de l'onglet règles
  RB_MODEGENE1 := TRadioButton(GetControl('RB_MODEGENE1'));
  RB_MODEGENE2 := TRadioButton(GetControl('RB_MODEGENE2'));
  RB_MODEGENE1.OnClick := fRB_MODEGENE1OnClick;
  RB_MODEGENE2.OnClick := fRB_MODEGENE2OnClick;

  RB_QUOTIDIENNE := TRadioButton(GetControl('RB_QUOTIDIENNE'));
  RB_NBINTERVENTION := TRadioButton(GetControl('RB_NBINTERVENTION'));
  RB_HEBDOMADAIRE := TRadioButton(GetControl('RB_HEBDOMADAIRE'));
  RB_ANNUELLE := TRadioButton(GetControl('RB_ANNUELLE'));
  RB_MENSUELLE := TRadioButton(GetControl('RB_MENSUELLE'));

  RB_QUOTIDIENNE.OnClick := fRB_QUOTIDIENNEOnClick;
  RB_NBINTERVENTION.OnClick := fRB_NBINTERVENTIONOnClick;
  RB_HEBDOMADAIRE.OnClick := fRB_HEBDOMADAIREOnClick;
  RB_ANNUELLE.OnClick := fRB_ANNUELLEOnClick;
  RB_MENSUELLE.OnClick := fRB_MENSUELLEOnClick;

  PC_FREQUENCE := TPageControl(GetControl('PC_FREQUENCE'));
  TS_QUOTIDIENNE := TTabSheet(GetControl('TS_QUOTIDIENNE'));
  TS_NBINTERVENTION := TTabSheet(GetControl('TS_NBINTERVENTION'));
  TS_HEBDOMADAIRE := TTabSheet(GetControl('TS_HEBDOMADAIRE'));
  TS_ANNUELLE := TTabSheet(GetControl('TS_ANNUELLE'));
  TS_MENSUELLE := TTabSheet(GetControl('TS_MENSUELLE'));

  TS_QUOTIDIENNE.TabVisible := false;
  TS_HEBDOMADAIRE.TabVisible := false;
  TS_NBINTERVENTION.TabVisible := false;
  TS_ANNUELLE.TabVisible := false;
  TS_MENSUELLE.TabVisible := false;

  RB_MOISMETHODEJOUR := TRadioButton(GetControl('RB_MOISMETHODEJOUR'));
  RB_MOISMETHODELE := TRadioButton(GetControl('RB_MOISMETHODELE'));
  RB_MOISMETHODETOUTLES := TRadioButton(GetControl('RB_MOISMETHODETOUTLES'));
  //
  SetControlProperty ('AFM_TYPEARTICLE', 'Plus', PlusTypeArticle (true));

  SetcontrolProperty('PCOMPLEMENT', 'Visible', False);

end;

procedure Tof_BTMODELETACHE.ModifZoneExit(Sender: TObject);
Begin

  // pour que la sortie de la zone permette la modification des règles
  fBoModifRegles := True;

end;


procedure Tof_BTMODELETACHE.BInsertOnClick(Sendre : TObject);
Begin
end;

procedure Tof_BTMODELETACHE.BDupliqueOnClick(Sendre : TObject);
Begin
end;

//Procedure de chargement des paramètre par défaut
procedure TOF_BTMODELETACHE.InitNewModele;
var vStTypeArticle: String;
    vStArticle    : String;
    vStCodeArticle: String;
    vStFacturable : String;
    vStUnite      : String;
    vStLibelle    : String;
    vStStatutPla  : String;
    frequence     : String;
    PR            : Double;
    PV            : Double;
Begin

  // article par défault
  //SetControlText('AFM_TYPEARTICLE', GetParamSocSecur('SO_BTTYPEARTDEF','PRE'));
  //SetControlText('AFM_CODEARTICLE', GetParamSocSecur('SO_BTPRESDEFAUT',''));

  vStCodeArticle := GetParamSocSecur('SO_BTPRESDEFAUT', '');
  vStTypeArticle := GetParamSocSecur('SO_BTTYPEARTDEF', 'PRE');
  vStFacturable  := '';
  vStUnite       := '';

  if ControleCodeArticle (vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle, PR, PV, vStStatutPla) then
     begin
     SetControlText('AFM_TYPEARTICLE', vStTypeArticle);
     SetControlText('AFM_CODEARTICLE', vStCodeArticle);
     UniteTemps.Value     := vStUnite;
     LibUnitTemps.caption := UniteTemps.Text;
     if GetControlText('AFM_LIBELLETACHE1') = '' then SetControlText('AFM_LIBELLETACHE1', vStLibelle);
     end
  else
     begin
     SetControlText('AFM_CODEARTICLE', '');
     SetControlText('AFM_TYPEARTICLE', GetParamsocSecur ('SO_BTTYPEARTDEF', 'PRE'));
     SetControlText('AFM_ARTICLE', '');
     end;

  SetControlEnabled('AFM_TYPEARTICLE', false);
  SetControlText('AFM_FAMILLETACHE', GetParamSocSecur('SO_BTFAMILLEDEF', ''));

  SetControlText('AFM_BTETAT', GetParamSocSecur('SO_BTETAT', ''));

  SetControlText ('JOURAN', '01');
  SetControlText ('MOISAN', '01');

  SetControlText('AFM_NBJOURSDECAL', GetParamSocSecur('SO_BTJOURDECAL', '5'));

  SetControlText ('AFM_DATEANNUELLE', DateToStr(now));
  SetControlText('AFM_MODECALC', '');

  setControlText ('AFM_DATEDEBPERIOD', DateToStr(now));
  setControlText ('AFM_DATEFINPERIOD', DateToStr(idate2099));

  setControlText ('AFM_MOISSEMAINE', '1');
  setControlText ('AFM_MODEGENE', '1');
  setControlText ('AFM_QTEINTERVENT', '1');

  Frequence := GetParamSocSecur('SO_BTPERIODDEFAUT', 'A');

  if Frequence = 'Q' then RB_QUOTIDIENNE.Checked := true;
  if Frequence = 'H' then RB_HEBDOMADAIRE.checked := True;
  if Frequence = 'M' then RB_MENSUELLE.Checked := True;
  if Frequence = 'A' then RB_ANNUELLE.Checked := true;
  if Frequence = 'N' then RB_NBINTERVENTION.Checked := true;

  setControlText ('AFM_NBJOURSDECAL', GetParamsocSecur ('SO_BTJOURSDECAL', 5));
  setControlText ('AFM_METHODEDECAL', GetParamsocSecur ('SO_BTJOURSPLANIF', 'P'));

end;

procedure TOF_BTMODELETACHE.InitNewRegles;
var
	fStPeriod : String;

begin
  // gestion des règles
  RB_MODEGENE1.checked := true;

  fStPeriod := GetParamsocSecur('SO_BTPERIODDEFAUT', 'Q');

  // C.B 23/05/2006
	// initialisation sur paramsoc
  if fStPeriod = 'A' then
		RB_ANNUELLE.checked := true

  else if fStPeriod = 'M' then
	  RB_MENSUELLE.checked := true

  else if fStPeriod = 'S' then
		RB_HEBDOMADAIRE.checked := true

  else if fStPeriod = 'Q' then
		RB_QUOTIDIENNE.checked := true

  else if fStPeriod = 'NBI' then
		RB_NBINTERVENTION.checked := true

  else if fStPeriod = '' then
	  RB_MENSUELLE.checked := true;

  RB_MOISMETHODEJOUR.checked := true;

  THValComboBox(GetControl('AFM_MOISSEMAINE1')).Value := '1';
  THValComboBox(GetControl('AFM_MOISJOURLIB')).Value := 'J1';
  THValComboBox(GetControl('AFM_MOISSEMAINE2')).Value := '1';

  SetCOntrolText('JOURAN', '1');
  SetCOntrolText('MOISAN', '1');
  SetControlText('AFM_MOISJOURFIXE', '1');
  SetControlText('AFM_MOISFIXE0', '1');
  SetControlText('AFM_MOISFIXE1', '1');
  SetControlText('AFM_MOISFIXE2', '1');

  TCheckBox(GetControl('AFM_JOUR1H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR2H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR3H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR4H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR5H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR6H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR7H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR1M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR2M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR3M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR4M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR5M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR6M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR7M')).checked := false;

end;


procedure TOF_BTMODELETACHE.fRB_MODEGENE1OnClick (Sender: TObject);
begin

  //RB_MODEGENE2.checked := not RB_MODEGENE1.Checked;
  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_MODEGENE2OnClick (Sender: TObject);
begin

  //RB_MODEGENE1.Checked := not RB_MODEGENE2.Checked;
  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_MOISMETHODEJOUROnClick(SEnder: TObject);
begin

  if RB_MOISMETHODEJOUR.Checked then
  begin
    SetControlEnabled('AFM_MOISJOURFIXE', True);
    SetControlEnabled('AFM_MOISFIXE0', True);
    SetControlEnabled('AFM_MOISSEMAINE1', false);
    SetControlEnabled('AFM_MOISJOURLIB', false);
    SetControlEnabled('AFM_MOISFIXE1', false);
    SetControlEnabled('AFM_MOISSEMAINE2', false);
    SetControlEnabled('AFM_MOISFIXE2', false);

    SetControlEnabled('AFM_JOUR1M', false);
    SetControlEnabled('AFM_JOUR2M', false);
    SetControlEnabled('AFM_JOUR3M', false);
    SetControlEnabled('AFM_JOUR4M', false);
    SetControlEnabled('AFM_JOUR5M', false);
    SetControlEnabled('AFM_JOUR6M', false);
    SetControlEnabled('AFM_JOUR7M', false);

    fBoModifRegles := true;
  end;

end;

procedure TOF_BTMODELETACHE.fRB_MOISMETHODELEOnClick(SEnder: TObject);
begin

  if RB_MOISMETHODELE.Checked then
  begin

    SetControlEnabled('AFM_MOISJOURFIXE', false);
    SetControlEnabled('AFM_MOISFIXE0', false);
    SetControlEnabled('AFM_MOISSEMAINE1', true);
    SetControlEnabled('AFM_MOISJOURLIB', true);
    SetControlEnabled('AFM_MOISFIXE1', true);
    SetControlEnabled('AFM_MOISSEMAINE2', false);
    SetControlEnabled('AFM_MOISFIXE2', false);

    SetControlEnabled('AFM_JOUR1M', false);
    SetControlEnabled('AFM_JOUR2M', false);
    SetControlEnabled('AFM_JOUR3M', false);
    SetControlEnabled('AFM_JOUR4M', false);
    SetControlEnabled('AFM_JOUR5M', false);
    SetControlEnabled('AFM_JOUR6M', false);
    SetControlEnabled('AFM_JOUR7M', false);

    fBoModifRegles := true;
  end;

end;

procedure TOF_BTMODELETACHE.fRB_MOISMETHODETOUTLESOnClick(SEnder: TObject);
begin

  if RB_MOISMETHODETOUTLES.Checked then
  begin

    SetControlEnabled('AFM_MOISJOURFIXE', false);
    SetControlEnabled('AFM_MOISFIXE0',    false);
    SetControlEnabled('AFM_MOISSEMAINE1', false);
    SetControlEnabled('AFM_MOISJOURLIB',  false);
    SetControlEnabled('AFM_MOISFIXE1',    false);
    SetControlEnabled('AFM_MOISSEMAINE2', true);
    SetControlEnabled('AFM_MOISFIXE2',    true);

    SetControlEnabled('AFM_JOUR1M', true);
    SetControlEnabled('AFM_JOUR2M', true);
    SetControlEnabled('AFM_JOUR3M', true);
    SetControlEnabled('AFM_JOUR4M', true);
    SetControlEnabled('AFM_JOUR5M', true);
    SetControlEnabled('AFM_JOUR6M', true);
    SetControlEnabled('AFM_JOUR7M', true);

    fBoModifRegles := true;
  end;

end;

procedure TOF_BTMODELETACHE.fRB_HEBDOMADAIREOnClick(Sender: TObject);
begin

  //Chargement de la Tob principale
  //fTobModele.PutValue('AFM_PLANNINGPERIOD', 'S');

  RB_QUOTIDIENNE.checked    := not RB_HEBDOMADAIRE.checked;
  RB_MENSUELLE.checked      := not RB_HEBDOMADAIRE.checked;
  RB_ANNUELLE.checked       := not RB_HEBDOMADAIRE.checked;
  RB_NBINTERVENTION.checked := not RB_HEBDOMADAIRE.checked;

  TCheckBox(GetControl('AFM_JOUR1H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR2H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR3H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR4H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR5H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR6H')).checked := not RB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('AFM_JOUR7H')).checked := not RB_HEBDOMADAIRE.checked;

  PC_FREQUENCE.ActivePage := TS_HEBDOMADAIRE;

  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_QUOTIDIENNEOnClick(Sender: TObject);
begin

  //Chargement de la Tob principale

  RB_HEBDOMADAIRE.checked   := not RB_QUOTIDIENNE.checked;
  RB_MENSUELLE.checked      := not RB_QUOTIDIENNE.checked;
  RB_ANNUELLE.checked       := not RB_QUOTIDIENNE.checked;
  RB_NBINTERVENTION.checked := not RB_QUOTIDIENNE.checked;

  PC_FREQUENCE.ActivePage   := TS_QUOTIDIENNE;

  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_MENSUELLEOnClick(Sender: TObject);
begin

  //Chargement de la Tob principale
  //fTobModele.PutValue('AFM_PLANNINGPERIOD', 'M');

  RB_HEBDOMADAIRE.checked   := not RB_MENSUELLE.checked;
  RB_QUOTIDIENNE.checked    := not RB_MENSUELLE.checked;
  RB_ANNUELLE.checked       := not RB_MENSUELLE.checked;
  RB_NBINTERVENTION.checked := not RB_MENSUELLE.checked;
  PC_FREQUENCE.ActivePage   := TS_MENSUELLE;

  TCheckBox(GetControl('AFM_JOUR1M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR2M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR3M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR4M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR5M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR6M')).checked := not RB_MENSUELLE.checked;
  TCheckBox(GetControl('AFM_JOUR7M')).checked := not RB_MENSUELLE.checked;

  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_ANNUELLEOnClick(Sender: TObject);
begin

  //Chargement de la Tob principale
  //fTobModele.PutValue('AFM_PLANNINGPERIOD', 'A');

  RB_HEBDOMADAIRE.checked   := not RB_ANNUELLE.checked;
  RB_QUOTIDIENNE.checked    := not RB_ANNUELLE.checked;
  RB_MENSUELLE.checked      := not RB_ANNUELLE.checked;
  RB_NBINTERVENTION.checked := not RB_ANNUELLE.checked;

  PC_FREQUENCE.ActivePage := TS_ANNUELLE;

  SetControlEnabled('RB_MODEGENE2', False);
  SetControlEnabled('RB_MODEGENE1', False);

  fBoModifRegles := true;

end;

procedure TOF_BTMODELETACHE.fRB_NBINTERVENTIONOnClick (Sender: TObject);
begin

  //Chargement de la Tob principale
  //fTobModele.PutValue('AFM_PLANNINGPERIOD', 'NBI');

  RB_HEBDOMADAIRE.checked := not RB_NBINTERVENTION.checked;
  RB_QUOTIDIENNE.checked  := not RB_NBINTERVENTION.checked;
  RB_MENSUELLE.checked    := not RB_NBINTERVENTION.checked;
  RB_ANNUELLE.checked     := not RB_NBINTERVENTION.checked;
  PC_FREQUENCE.ActivePage := TS_NBINTERVENTION;

  SetControlEnabled ('RB_MODEGENE2', false);
  SetControlEnabled ('RB_MODEGENE1', false);

  fBoModifRegles := true;

end;


Procedure TOF_BTMODELETACHE.ArticleOnexit(sender : Tobject);
var vStCodeArticle  : String;
    vStTypeArticle  : String;
    vStArticle      : String;
    vStFacturable   : String;
    vStUnite        : String;
    VstLibelle      : string;
    vStStatutPla    : String;
    PR              : Double;
    PV              : double;
begin

  vStCodeArticle := GetControlText('AFM_CODEARTICLE');
  if vStCodeArticle = '' then Exit;
  if ControleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStunite, vStLibelle, PR, PV, vStStatutPla) then
     begin
     //SetControlText ('AFM_TYPEARTICLE', vStTypeArticle);
     if IsTachePrestation(vStTypeArticle) then    //AB-200605 - FQ 12918
        SetControlText('AFM_UNITETEMPS', vStUnite)
     else
        SetControlText('AFM_UNITETEMPS', 'J'); // pour frais et marchandise on met à J
        if getControlText('AFM_LIBELLETACHE1') = '' then
           SetControlText('AFM_LIBELLETACHE1', vStLibelle);
     end
  else
     begin
     PGIBoxAF (TexteMessage [3] , '');    {3, 'L''article n''existe pas.'}
     exit;
     end;

  //if getControlText ('AFM_CODEARTICLE') <> '' then
  //   SetControlEnabled ('AFM_TYPEARTICLE', False)
  //else
  //  SetControlEnabled ('AFM_TYPEARTICLE', true);

end;


Procedure TOF_BTMODELETACHE.ArticleOnElipsisClick(sender : Tobject);
begin

  if GetControlText('AFM_LIBELLETACHE1') = fStLibelleParDefaut then SetControlText('AFM_LIBELLETACHE1', '');

  SetControlText('AFM_CODEARTICLE', DispatchRecherche(nil, 1, '', 'GA_CODEARTICLE='
+ GetControltext('AFM_CODEARTICLE') + ';RETOUR_CODEARTICLE=X;GA_TYPEARTICLE='
+ GetControlText('AFM_TYPEARTICLE'), ''));

  RaffraichirArticle;

end;

procedure TOF_BTMODELETACHE.RaffraichirArticle;
var vStCodeArticle  : string;
    vStTypeArticle  : string;
    vStArticle      : string;
    vStFacturable   : string;
    vStUnite        : string;
    vStLibelle      : string;
    vStStatutPla    : String;
    PR,PV           : Double;
begin

  vStCodeArticle := GetControlText('AFM_CODEARTICLE');

  if vStCodeArticle <> '' then
     begin
     if ControleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle, PR, PV, vStStatutPla) then
        begin
        SetControlText('AFM_TYPEARTICLE', vStTypeArticle);
        if IsTachePrestation (vStTypeArticle) then  //AB-200605 - FQ 12918
           begin
           SetControlText('AFM_UNITETEMPS', vStUnite);
           end
        else
           SetControlText('AFM_UNITETEMPS', 'J'); // pour frais et marchandise on met à J
        end;
     SetControlText('AFM_LIBELLETACHE1', vStLibelle);
     SetControlText('AFM_ARTICLE', vStArticle);
     end;

end;

procedure TOF_BTMODELETACHE.MODELETACHEOnexit(Sender: TObject);
var Req  : String;
    QQ   : TQuery;
begin

  fStatut := taModif;

  if fAction = taModif then exit;

  //Lecture de la table modele de tache pour
  //vérification existance code ou non en création
  CodeModele := GetControlText('AFM_MODELETACHE');

  if CodeModele = '' then exit;

  req := 'SELECT * FROM AFMODELETACHE WHERE AFM_MODELETACHE ="' + CodeModele + '"';
  QQ := OpenSql(Req,true);

  if QQ.eof then
     begin
     ferme(QQ);
     Exit;
     end;

  fAction := taModif;

  fTobModele.selectDB('',QQ);
  ferme (QQ);

  onload;

end;

procedure TOF_BTMODELETACHE.TypeActionOnExit(Sender: TObject);
var StReq : String;
    QQ    : TQuery;
Begin

  fStatut := taModif;

  if Pos(GetControlText('AFM_BTETAT'),'ABSPAIE;ACT-GRC')>0 then
  begin
  	// type réservé
     PGIBoxAF (TexteMessage [10] , '');    // réservé
     THEdit(GetControl('AFM_BTETAT')).SetFocus;
     exit;
  end;
  LectureTypeAction(GetControlText('AFM_BTETAT'));

  if GetControlText('AFM_BTETAT') = '' then exit;

  StReq := 'SELECT BTA_DUREEMINI FROM BTETAT WHERE BTA_BTETAT = "' + GetControlText('AFM_BTETAT') + '"';
  QQ := nil;

  try
    QQ := OpenSQL(StReq, True);
    if (not QQ.EOF) then
       begin
       setControlText('AFM_QTEINTERVENT', FloatToStr(QQ.findField('BTA_DUREEMINI').asfloat));
       end;
  finally
    ferme(QQ);
  end;

end;

procedure TOF_BTMODELETACHE.LectureTypeAction (TypeAction : string);
var Req           : String;
    TobTypeAction : Tob;
begin

  if TypeAction = '' then exit;

  SetControlText('LTYPEACTION', '');

  //Lecture du Type ACtion Evenement sélectionné et affichage des informations
  req := 'SELECT BTA_LIBELLE FROM BTETAT WHERE BTA_BTETAT="' + TypeAction + '"';

  TobTypeAction := Tob.Create('#BTETAT',Nil, -1);
  TobTypeAction.LoadDetailFromSQL(Req);
  if TobTypeAction.Detail.Count > 0 then
     Setcontroltext('LTYPEACTION', TobTypeAction.detail[0].GetString('BTA_LIBELLE'))
  else
     Setcontroltext('LTYPEACTION', '');

  TobTypeAction.free;

end;

procedure TOF_BTMODELETACHE.fED_DATEPERIODDEBOnExit(SEnder: TObject);
var DateDeb : TDateTime;
begin

  if TFFiche (Ecran).TypeAction = taConsult then exit;

  fStatut := taModif;

  DateDeb :=StrToDateTime(GetControlText ('AFM_DATEDEBPERIOD'));

  FtobModele.Putvalue('AFM_DATEDEBPERIOD', DateDeb);

end;

procedure TOF_BTMODELETACHE.fED_DATEPERIODFINOnExit(SEnder: TObject);
var DateFin : TDateTime;
begin

  if TFFiche (Ecran).TypeAction = taConsult then  exit;

  fstatut := taModif;

  DateFin :=StrToDateTime(GetControlText ('AFM_DATEFINPERIOD'));

  FtobModele.Putvalue('AFM_DATEFINPERIOD', DateFin);

end;

procedure TOF_BTMODELETACHE.OnClose;
begin
  inherited;

  if fTobModele <> nil then fTobModele.free

end;

procedure TOF_BTMODELETACHE.AFM_UNITETEMPSOnChange(Sender: TObject);
begin

  LibUnitTemps.Caption := UniteTemps.text;
  
end;

initialization
  registerclasses ([TOF_BTMODELETACHE] );
end.
