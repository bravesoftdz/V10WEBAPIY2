{***********UNITE*************************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 01/03/2006
Description .. : TOM commune à tous les métiers
Mots clefs ... : uTOMComm

TOM            : TOM AGL
 |
 +- tTOMComm   : TOM Commune à tous les métiers
     |           -
     |           -
     |
     +- twTOM  : TOM Métier W (Gestion de Production)
*****************************************************************}

Unit uTOMComm;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  Variants,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}
      dbtables,
    {$ELSE DBEXPRESS}
      uDbxDataSet,
    {$ENDIF DBEXPRESS}
    {$IFNDEF EAGLSERVER}
      Fiche,
      FichList,
    {$ENDIF !EAGLSERVER}
  {$ELSE !EAGLCLIENT}
    eFiche,
    eFichList,
  {$ENDIF !EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      SaisieList,
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob
  ;

Type
  tScriptEvent = (SEOnNewRecord, SEOnDeleteRecord, SEOnAfterDeleteRecord, SEOnUpdateRecord, SEOnAfterUpdateRecord);
  tTOMComm = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    fAfterUpdating : boolean;   { Je viens du OnAfterUpdate }
    fDoControl     : Integer;   { Dois-je faire un controle sur le champ? }
    fMemoIkc       : string;    { Mémorisation du ikc pour le OnAfterUpdate }
    fStArgument    : String;    { Mémorisation de S de argument }
    {$IFNDEF EAGLCLIENT}
      MyfTOB       : Tob;
    {$ENDIF EAGLCLIENT}

	procedure GetRetour;
    function GetIkc: string;

    {$IFNDEF EAGLSERVER}
      {$IFNDEF PGIMAJVER}
        {$IFNDEF ERADIO}
          { Loupe }
          procedure PmLoupe_OnPopUp(Sender: TObject);
          procedure MnProperties_OnClick(Sender: TObject);
          procedure SetLibelleLibre;
          procedure SetLibelleFamille;
        {$ENDIF !ERADIO}
      {$ENDIF !PGIMAJVER}
      { Script GEP }
{$IFDEF GCGC}
      procedure LanceScript(Const ScriptEvent: tScriptEvent);
{$ENDIF GCGC}
      { Générique }
      procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    {$ENDIF !EAGLSERVER}
//GP_20080208_TS_GP14792
    procedure CheckIOError;
  protected
    function GetLoupeCtx: string; Virtual;
    procedure ControlField(Const FieldName: string); virtual;
    procedure CalculField(Const FieldName: string); virtual;
    function RecordIsvalid: boolean; virtual;
  public
    { Control écran }
    procedure DisableControl;
    procedure EnableControl;
    function CanControl: Boolean;
    procedure SetControlsVisible(Const FieldsName: Array of String; Const Visibility: Boolean);
    procedure SetControlsEnabled(Const FieldsName: Array of String; Const Enability: Boolean);

    function EcranIs(const FormName: String): Boolean;

    { Get }
    function GetInteger(Const FieldName: string): integer;
    function GetString(Const FieldName: string): string;
    function GetDateTime(Const FieldName: string): tDateTime;
    function GetDouble(Const FieldName: string): Double;
    function GetBoolean(Const FieldName: string): Boolean;

    function GetIntegerAvantModif(Const FieldName: string): integer;
    function GetStringAvantModif(Const FieldName: string): String;
    function GetDateTimeAvantModif(Const FieldName: string): tDateTime;
    function GetDoubleAvantModif(Const FieldName: string): Double;
    function GetBooleanAvantModif(Const FieldName: string): Boolean;

    function GetTableName: string;
    function GetPrefixe: string;
    function FieldExists(const FieldName: string): boolean;
    function GetTob: Tob;
    {$IF not(Defined(EAGLSERVER) or Defined(ERADIO))}
    function GetTobAlertes: Tob;
    {$IFEND !EAGLSERVER && !ERADIO}

    { Set }
    procedure SetString(Const FieldName: string; Const Value: String; Const WithControlField: boolean = false);
    procedure SetInteger(Const FieldName: string; Const Value: integer; Const WithControlField: boolean = false);
    procedure SetDateTime(Const FieldName: string; Const Value: tDateTime; Const WithControlField: boolean = false);
    procedure SetDouble(Const FieldName: string; Const Value: double; Const WithControlField: boolean = false);
    procedure SetBoolean(Const FieldName: string; Const Value: boolean; Const WithControlField: boolean = false);

    procedure SetEtablissement;
    procedure SetDomaine;
    procedure SetDepot;

    procedure SetBlocNote;
    procedure wSetBmemo;
    procedure SetGuid;

    property Ikc: string read GetIkc;                               { Ikc: 'C', 'M', 'S' }
    property StArgument: String read FStArgument Write FStArgument;
  end;

//GP_20080208_TS_GP14792 >>>
const
  ErrTransactionTom = 1001;
//GP_20080208_TS_GP14792 <<<

Implementation

uses
  Math,
  wCommuns,
  Menus,
  UtilPGI,     //js1 pour prendre en compte transfert de isoracle wcommuns==>utilpgi
  {$IFNDEF PGIMAJVER}
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        wMnu,
      {$ENDIF !ERADIO}
{$IFDEF GCGC}
      GcOuvertureGEP,
      UtilGC,
{$ENDIF GCGC}
    {$ENDIF !EAGLSERVER}
  {$ENDIF !PGIMAJVER}
  Ent1
{$IFDEF GRC}
  ,EntRT,ParamSoc
{$ENDIF GRC}    
  ;

procedure tTOMComm.OnNewRecord ;
begin
  FAfterUpdating := False;
  Inherited ;

  { Domaine;Depôt;Etablissement par défaut }
  SetDomaine;
  SetDepot;
  SetEtablissement;

  {$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
    { Script - GEP }
    LanceScript(SEOnNewRecord);
{$ENDIF GCGC}
  {$ENDIF !EAGLSERVER}
end ;

procedure tTOMComm.OnDeleteRecord ;
begin
  FAfterUpdating := False;

  {$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
    { Script - GEP }
    LanceScript(SEOnDeleteRecord);
{$ENDIF GCGC}
  {$ENDIF !EAGLSERVER}

  Inherited ;

//GP_20080208_TS_GP14792
  CheckIOError()
end ;

procedure tTOMComm.OnUpdateRecord ;
begin
  fMemoIkc := Ikc;
  FAfterUpdating := False;

  Inherited ;

  DisableControl;
  try
    if LastError = 0 then
    begin
      SetBlocNote;
      wSetBMemo;
      SetGuid;
      {$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
        { Script - GEP }
        LanceScript(SEOnUpdateRecord);
{$ENDIF GCGC}
      {$ENDIF !EAGLSERVER}

      { Retourne la clé actuelle de la fiche pour 'refresher' le Mul }
      GetRetour;

    end
  finally
    EnableControl;
    fDoControl := 0;
  end;
//GP_20080208_TS_GP14792
  CheckIOError()
end;

procedure tTOMComm.OnAfterUpdateRecord ;
begin
  FAfterUpdating := True;
  Inherited ;
  {$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
    { Script - GEP }
    LanceScript(SEOnAfterUpdateRecord);
{$ENDIF GCGC}
  {$ENDIF !EAGLSERVER}

//GP_20080208_TS_GP14792
  CheckIOError()
end ;

procedure tTOMComm.OnAfterDeleteRecord ;
begin
  {$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
    { Script - GEP }
    LanceScript(SEOnAfterDeleteRecord);
{$ENDIF GCGC}
  {$ENDIF !EAGLSERVER}

  Inherited ;

//GP_20080208_TS_GP14792
  CheckIOError()
end ;

procedure tTOMComm.OnLoadRecord ;
begin
  FAfterUpdating := False;
  Inherited ;

  if Assigned(GetControl('MNPROPERTIES')) then
    SetControlEnabled('MNPROPERTIES', Ikc <> 'C');

  if Ikc <> 'C' then
  begin
    GetRetour;
  end;

end ;

procedure tTOMComm.OnChangeField ( F: TField ) ;
var
  FieldName: string;
begin
  FieldName := F.FieldName; { Mémorisé car change en CWas... }
  
  Inherited ;

  if Assigned(DS) then
  begin
    if (DS.State <> dsBrowse) and (fDoControl = 0) and (not AfterInserting) then
    begin
      ControlField(FieldName);
    end;
  end;

  CalculField(FieldName);
end ;

procedure tTOMComm.OnArgument ( S: String ) ;
begin
  StArgument := S;

  {$IFNDEF EAGLCLIENT}
    { Tob des données }
    if Assigned(Ecran) then
      MyfTob := Tob.Create(GetTableName, nil, -1);
  {$ENDIF EAGLCLIENT}

  Inherited ;

  fDoControl := 0;

  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      {$IFNDEF PGIMAJVER}
        { Libellé libres }
        SetLibelleLibre;

        { Libellés des familles }
        SetLibelleFamille;

        if Assigned(GetControl('PMLOUPE')) then
          TPopUpMenu(GetControl('PMLOUPE')).OnPopUp := PmLoupe_OnPopUp;
        if Assigned(GetControl('MNPROPERTIES')) then
          TMenuItem(GetControl('MNPROPERTIES')).OnClick := MnProperties_OnClick;
      {$ENDIF !PGIMAJVER}
    {$ENDIF !ERADIO}

    { Raccourci clavier}
    if Assigned(Ecran) then
      Ecran.OnKeyDown := FormKeyDown;
  {$ENDIF !EAGLSERVER}

  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  if Assigned(GetControl('MNDISPO')) then
    TMenuItem(GetControl('MNDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).ShortCut := TextToShortCut('Ctrl+A');
end ;

procedure tTOMComm.OnClose ;
begin
  Inherited ;

  {$IFNDEF EAGLCLIENT}
    { Tob des données }
    if Assigned(MyfTob) then
      FreeAndNil(MyfTob);
  {$ENDIF EAGLCLIENT}
end ;

procedure tTOMComm.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure tTOMComm.SetBoolean(const FieldName: string; const Value, WithControlField: boolean);
begin
  SetField(FieldName, BoolToStr_(Value));
  if WithControlField then
    ControlField(FieldName);
end;

procedure tTOMComm.SetDateTime(const FieldName: string; const Value: tDateTime; const WithControlField: boolean);
begin
  SetField(FieldName, Value);
  if WithControlField then
    ControlField(FieldName);
end;

procedure tTOMComm.SetDouble(const FieldName: string; const Value: double; const WithControlField: boolean);
begin
  SetField(FieldName, Value);
  if WithControlField then
    ControlField(FieldName);
end;

procedure tTOMComm.SetInteger(const FieldName: string; const Value: integer; const WithControlField: boolean);
begin
  SetField(FieldName,Value);
  if WithControlField then
    ControlField(FieldName);
end;

procedure tTOMComm.SetString(const FieldName, Value: String; const WithControlField: boolean);
begin
  SetField(FieldName,Value);
  if WithControlField then
    ControlField(FieldName);
end;

procedure tTOMComm.ControlField(const FieldName: string);
begin
  if (LastError <> 0) and Assigned(Ecran) then
    SetFocusControl(FieldName);
end;

procedure tTOMComm.DisableControl;
begin
  Inc(fDoControl);
end;

procedure tTOMComm.EnableControl;
begin
  fDoControl := Max(0, fDoControl-1);
end;

procedure tTOMComm.CalculField(const FieldName: string);
begin
{}
end;

function tTOMComm.RecordIsvalid: boolean;
begin
  Result := true;
end;

Procedure tTOMComm.GetRetour;
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    if (Assigned(Ecran)) and (Ecran is TFFiche) and (TFFICHE(ECRAN).Retour = '') then
      TFFICHE(ECRAN).Retour := wGetValueClef1(GetTableName, TFFICHE(ECRAN));
  {$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
end;

function tTOMComm.GetIkc: string;
begin
  if Assigned(Ecran) and Assigned(DS) then
  begin
    if FAfterUpdating then
      Result := FMemoIkc
    else
      Result := iif(DS.State = dsInsert, 'C', 'M')
  end
 	else
    Result := GetString('IKC');
end;

function tTOMComm.GetBoolean(const FieldName: string): Boolean;
begin
	if GetField(FieldName) <> null then
    Result := StrToBool_(GetField(FieldName))
  else
    Result := False;
end;

function tTOMComm.GetDateTime(const FieldName: string): tDateTime;
begin
	if GetField(FieldName) <> null then
    Result := GetField(FieldName)
  else
    Result := iDate1900;
end;

function tTOMComm.GetDouble(const FieldName: string): Double;
begin
	if GetField(FieldName) <> null then
    Result := GetField(FieldName)
  else
    Result := 0.0;
end;

function tTOMComm.GetInteger(const FieldName: string): integer;
begin
	if GetField(FieldName) <> null then
    Result := GetField(FieldName)
  else
    Result := 0;
end;

function tTOMComm.GetString(const FieldName: string): string;
begin
  if VarToStr(GetField(FieldName)) <> #0 then
    Result := VarToStr(GetField(FieldName))
  else
    Result := ''
end;

function tTOMComm.GetBooleanAvantModif(const FieldName: string): Boolean;
begin
	if GetFieldAvantModif(FieldName) <> null then
    Result := StrToBool_(GetFieldAvantModif(FieldName))
  else
    Result := False;
end;

function tTOMComm.GetDateTimeAvantModif(const FieldName: string): tDateTime;
begin
	if GetFieldAvantModif(FieldName) <> null then
    Result := GetFieldAvantModif(FieldName)
  else
    Result := iDate1900;
end;

function tTOMComm.GetDoubleAvantModif(const FieldName: string): Double;
begin
	if GetFieldAvantModif(FieldName) <> null then
    Result := GetFieldAvantModif(FieldName)
  else
    Result := 0.0;
end;

function tTOMComm.GetIntegerAvantModif(const FieldName: string): integer;
begin
	if GetFieldAvantModif(FieldName) <> null then
    Result := GetFieldAvantModif(FieldName)
  else
    Result := 0;
end;

function tTOMComm.GetStringAvantModif(const FieldName: string): String;
begin
	if GetFieldAvantModif(FieldName) <> null then
    Result := GetFieldAvantModif(FieldName)
  else
    Result := '';
end;

{$IFNDEF ERADIO}
{$IFNDEF PGIMAJVER}
{$IFNDEF EAGLSERVER}
procedure tTOMComm.MnProperties_OnClick(Sender: TObject);
var
  Identifiant, Clef: String;
begin
  if Ikc <> 'C' then
  begin
    Identifiant := wGetValueClef1(GetTableName, Ecran);
    if Identifiant <> '' then
    begin
      Identifiant := StringReplace(Identifiant, ';', '~', [rfIgnoreCase, rfReplaceAll]);
      Clef := wMakeFieldString(GetTableName, '~');

      wCallProperties(GetPrefixe, Identifiant, Clef, IntToStr(LongInt(Ecran)));
    end
    else
      PGIInfo(TraduireMemoire('L''enregistrement est vide.'), TraduireMemoire('Proprietés'));
  end
end;
{$ENDIF !EAGLSERVER}
{$ENDIF !PGIMAJVER}
{$ENDIF !ERADIO}

function tTOMComm.GetTableName: string;
begin
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  if Ecran is TFSAISIELIST then
    Result := UpperCase(TFSaisieList(Ecran).LeFiltre.TableName)
  else
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
    Result := TableName;
end;

function tTOMComm.GetPrefixe: string;
begin
  Result := TableToPrefixe(GetTableName);
end;

function tTOMComm.FieldExists(const FieldName: string): boolean;
begin
  if Assigned(fTob) then
    Result := fTob.FieldExists(FieldName)
  else
    Result := wExistFieldInDechamps(FieldName)
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tTOMComm.PmLoupe_OnPopUp(Sender: TObject);
begin
  wSetMnuLoupe(false, GetLoupeCtx, tPopUpMenu(GetControl('PMLOUPE')));
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}
{$ENDIF !EAGLSERVER}

function tTOMComm.GetLoupeCtx: string;
begin
  Result := '';

  if GetPrefixe = 'GA' then
  begin
    Result := 'TYPEARTICLE=' + GetString('GA_TYPEARTICLE');
  end
  ;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tTOMComm.SetLibelleLibre;
var
  iLibre                                    : integer;
  Suffixe, MasterPrefixe, FieldName, MasterFieldName,NaturePrefixe : string;

  procedure SetVisible(FieldName: string);
  begin
    SetControlProperty('T' + FieldName, 'VISIBLE', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
    SetControlProperty(FieldName, 'VISIBLE', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
  end;
  function GetNaturePrefixe: string;
  begin
    if Pos(UpperCase(Copy(MasterPrefixe, 1, 2)), 'QW;QO;QU;QJ;QK') > 0 then
      Result := 'QMES'
{$IFDEF QUALITE}
    else if UpperCase(Copy(MasterPrefixe, 1, 2)) = 'RQ' then
      Result := 'RQ'
{$ENDIF QUALITE}
    else if Copy(MasterPrefixe, 1, 2) = 'GC' then
      Result := 'GC'
    else
      Result := 'W'
  end;
begin
  { Exit }
  Suffixe := GetPrefixe;
  if Suffixe = 'QWL' then
    Suffixe := 'QWH';
  FieldName := GetPrefixe + '_LIBRE' + Suffixe + '1';
  if Assigned(GetControl(FieldName)) then
  begin
    {$IFDEF GPAOLIGHT}
    MasterPrefixe := GetMasterPrefixe(GetPrefixe);
    {$ELSE GPAOLIGHT}
    MasterPrefixe := GetPrefixe;
    {$ENDIF GPAOLIGHT}
    if GetPrefixe = 'QWL' then
      MasterPrefixe := 'QWH';

    { Libellés champs libres }
    for iLibre := 1 to 10 do
    begin
      { tables libres }
      if (iLibre<=3) or (Pos(MasterPrefixe , 'QWH;QWG;QWL') = 0) then
      begin
        FieldName := GetPrefixe + '_LIBRE' + Suffixe + intToHex(iLibre, 1);
        MasterFieldName := MasterPrefixe + '_LIBRE' + MasterPrefixe + intToHex(iLibre, 1);

        NaturePrefixe := GetNaturePrefixe;

        SetControlText('T' + FieldName,RechDom(NaturePrefixe + 'LIBELLELIBRE' + MasterPrefixe, 'T' + MasterFieldName, false));
        SetVisible(FieldName);
      end;
      if iLibre <= 3 then
      begin
        { Valeurs libres }
        FieldName := GetPrefixe + '_VALLIBRE' + intToStr(iLibre);
        MasterFieldName := MasterPrefixe + '_VALLIBRE' + intToStr(iLibre);
        SetControlProperty('T' + FieldName, 'CAPTION', RechDom(NaturePrefixe + 'VALLIBRE' + MasterPrefixe, 'T' + MasterFieldName, false));
        SetVisible(FieldName);

        { Dates libres }
        FieldName := GetPrefixe + '_DATELIBRE' + intToStr(iLibre);
        MasterFieldName := MasterPrefixe + '_DATELIBRE' + intToStr(iLibre);
        SetControlProperty('T' + FieldName, 'CAPTION', RechDom(NaturePrefixe+'DATELIBRE' + MasterPrefixe, 'T' + MasterFieldName, false));
        SetVisible(FieldName);

        { Décisions libres }
        FieldName := GetPrefixe + '_BOOLLIBRE' + intToStr(iLibre);
        MasterFieldName := MasterPrefixe + '_BOOLLIBRE' + intToStr(iLibre);
        SetControlProperty(FieldName, 'CAPTION', RechDom(NaturePrefixe + 'BOOLLIBRE' + MasterPrefixe, 'T' + MasterFieldName, false));
        SetControlProperty(FieldName, 'VISIBLE', Copy(RechDom(NaturePrefixe + 'BOOLLIBRE' + MasterPrefixe, 'T' + MasterFieldName, false), 1, 2) <> '.-');

        { Textes libres }
        if (Pos(MasterPrefixe , 'QWH;QWG;QWL') = 0) then
        begin
          FieldName :=  GetPrefixe + '_CHARLIBRE' + intToStr(iLibre);
          MasterFieldName := MasterPrefixe + '_CHARLIBRE' + intToStr(iLibre);
          SetControlProperty('T' + FieldName, 'CAPTION', RechDom(NaturePrefixe + 'CHARLIBRE' + MasterPrefixe, 'T' + MasterFieldName, false));
          SetVisible(FieldName);
        end;
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tTOMComm.SetLibelleFamille;
var
  Visible : Boolean;
  c       : Char;
  Caption : string;
begin
  for c := '1' to '3' do
  begin
    Caption := RechDom('GCLIBFAMILLE', 'LF' + c, False);
    Visible := Copy(Caption, 1, 2) <> '.-';
    if Assigned(GetControl('TGA_FAMILLENIV' + c)) then
      SetControlCaption('TGA_FAMILLENIV' + c, Caption);
    SetControlVisible('TGA_FAMILLENIV' + c, Visible);
    SetControlVisible('GA_FAMILLENIV'  + c, Visible);
    if Assigned(GetControl('T' + GetPrefixe + '_FAMILLENIV' + c)) then
      SetControlCaption('T' + GetPrefixe + '_FAMILLENIV' + c, Caption);
    SetControlVisible('T' + GetPrefixe + '_FAMILLENIV' + c, Visible);
    SetControlVisible(GetPrefixe + '_FAMILLENIV' + c, Visible);
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}
{$ENDIF !EAGLSERVER}

procedure tTOMComm.SetDepot;
begin
  if GetPrefixe = 'WPF' then exit;

  if FieldExists(GetPrefixe + '_DEPOT') then
  begin
    SetString(GetPrefixe + '_DEPOT', VH^.ProfilUserC[PrEtablissement].Depot);

    if Assigned(Ecran) and Assigned(GetControl(GetPrefixe + '_DEPOT')) and tWinControl(GetControl(GetPrefixe + '_DEPOT')).Enabled then
      SetControlEnabled(GetPrefixe + '_DEPOT', not VH^.ProfilUserC[PrEtablissement].ForceDepot);
  end;
end;

procedure tTOMComm.SetDomaine;
begin
  if GetPrefixe = 'WPF' then exit;

  if FieldExists(GetPrefixe + '_DOMAINE') then
  begin
    SetString(GetPrefixe + '_DOMAINE', VH^.ProfilUserC[PrEtablissement].Domaine);

    if Assigned(Ecran) and Assigned(GetControl(GetPrefixe + '_DOMAINE')) and tWinControl(GetControl(GetPrefixe + '_DOMAINE')).Enabled then
      SetControlEnabled(GetPrefixe + '_DOMAINE', not VH^.ProfilUserC[PrEtablissement].ForceDomaine);
  end;
end;

procedure tTOMComm.SetEtablissement;
begin
  if GetPrefixe = 'WPF' then exit;

  if FieldExists(GetPrefixe + '_ETABLISSEMENT') then
  begin
    SetString(GetPrefixe + '_ETABLISSEMENT', VH^.ProfilUserC[PrEtablissement].Etablissement);

    if Assigned(Ecran) and Assigned(GetControl(GetPrefixe + '_ETABLISSEMENT')) and tWinControl(GetControl(GetPrefixe + '_ETABLISSEMENT')).Enabled then
      SetControlEnabled(GetPrefixe + '_ETABLISSEMENT', not VH^.ProfilUserC[PrEtablissement].ForceEtab);
  end;
end;

procedure tTOMComm.SetBlocNote;
  procedure SetBlocNote(Const FieldName: string);
  begin
    if (ChampToNum(FieldName) > 0) and (GetString(FieldName) = '') then
      SetString(FieldName, #0)
  end;
begin
  if IsOracle then
  begin
    SetBlocNote(GetPrefixe + '_BLOCNOTE');
    SetBlocNote(GetPrefixe + '_BLOC_NOTE');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 03/04/2002
Modifié le ... :   /  /
Description .. : Mise à jour du champ WBMEMO à vrai si BLOCNOTE <> ''
Mots clefs ... :
*****************************************************************}
procedure tTOMComm.wSetBmemo;
var
  xLastError     : integer;
  xLastErrorMsg  : string;
begin
  if (ChampToNum(GetPrefixe + '_WBMEMO') > 0) and (ChampToNum(GetPrefixe + '_BLOCNOTE') > 0) then
  begin
    { Bien mémoriser le LastError avant }
    xLastError := LastError;
    xLastErrorMsg := LastErrorMsg;
    try
      SetBoolean(GetPrefixe + '_WBMEMO', (GetString(GetPrefixe + '_BLOCNOTE') <> ''));
    finally
      LastError := xLastError;
      LastErrorMsg := xLastErrorMsg;
    end;
  end;
end;

procedure tTOMComm.SetGuid;
begin
  if     (Ikc = 'C')
     and (FieldExists(GetPrefixe + '_GUID'))
     and ((not FieldExists('GUIDINITIALISE')) or (GetString(GetPrefixe + '_GUID') = '')) then
  begin
    SetString(GetPrefixe + '_GUID', AGLGetGuid);
  end;
end;

function tTOMComm.GetTob: Tob;
begin
  if not Assigned(Ecran) then
    Result := fTob
  else
  begin
    {$IFDEF EAGLCLIENT}
      Result:= DS.CurrentFille;
    {$ELSE EAGLCLIENT}
      wMakeTobFromDS(DS , MyfTob);
      Result := MyfTob;
    {$ENDIF EAGLCLIENT}
  end;
end;

{$IF not(Defined(EAGLSERVER) or Defined(ERADIO))}
function tTOMComm.GetTobAlertes: Tob;
var i : integer ;
    TOBInfosCompl,TobCl : tob;
begin
  if not Assigned(Ecran) then
    Result := fTob
  else
  begin
    Result := GetTob;
    {$IFDEF GRC}
      VH_RT.TobChampsPro.Load;

      TobCl:=VH_RT.TobParamCl.FindFirst(['CO_TYPE','CO_ABREGE'],['RPC',Ecran.name],TRUE) ;
      if Assigned(TobCl) then
        begin
        if GetParamSocSecur('SO_RTGESTINFOS00'+TobCl.GetString('CO_CODE'),false) then
          begin
          TOBInfosCompl:= TOB.Create('RTINFOS00'+TobCl.GetString('CO_CODE'), nil, -1);
          if not inserting then
            TOBInfosCompl.selectDB ('"'+wGetValueClef1(GetTableName, TFFICHE(ECRAN))+'"',nil);
            for i := 1 to Pred(TOBInfosCompl.NbChamps) do
              Result.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
          TOBInfosCompl.free;
          end;
        end;
    {$ENDIF GRC}
  end;
end;
{$IFEND !EAGLSERVER && !ERADIO}

{$IFDEF GCGC}
{$IFNDEF EAGLSERVER}
procedure tTOMComm.LanceScript(Const ScriptEvent: tScriptEvent);
{$IFNDEF PGIMAJVER}
var
  iChamp : integer;
  TobData: Tob;
{$ENDIF !PGIMAJVER}
begin
{$IFNDEF PGIMAJVER}
	if LastError <> 0 then exit;
  if V_Pgi.CGEPDisabled then exit;
  // Dans le cas "OnAfterDelete" les données on été supprimées, on se trouve sur l'enregistrement précedent -> on ne fait rien
  //  (sauf si appel hors écran, auquel cas la tob correspond au bon enregistrement)
  if (ScriptEvent=SEOnAfterDeleteRecord) and (Assigned(Ecran)) then exit;

  if Assigned(Ecran) then
  begin
    if GetTob <> nil then
    begin
      TobData := Tob.Create(GetTableName, nil, -1);
      TobData.Dupliquer(GetTob, false, true);
    end;  
  end
  else
  begin
    TobData := fTob;
  end;

  if Assigned(TobData) then
  begin
    try
      Case ScriptEvent of
        SEOnNewRecord   :       GcOuvertureGEP.OnNewRecord(GetTableName, TobData);
        SEOnDeleteRecord:       GcOuvertureGEP.OnDeleteRecord(GetTableName, TobData);
        SEOnAfterDeleteRecord:  GcOuvertureGEP.OnAfterDeleteRecord(GetTableName, TobData);
        SEOnUpdateRecord:       GcOuvertureGEP.OnUpdateRecord(GetTableName, Ikc, TobData);
        SEOnAfterUpdateRecord:  GcOuvertureGEP.OnAfterUpdateRecord(GetTableName, Ikc, TobData);
      else
        PgiError('Script non implémenté', 'tTOMComm.LanceScript')
      end;
    finally
      if Assigned(Ecran) then
      begin
        if not (ScriptEvent in [SEOnAfterDeleteRecord,SEOnAfterUpdateRecord]) then
        begin
          for iChamp := 1 to TobData.NbChamps do
          begin
            if GetField(TobData.GetNomChamp(iChamp))<> TobData.GetValue(TobData.GetNomChamp(iChamp)) then
              SetField(TobData.GetNomChamp(iChamp), TobData.GetValue(TobData.GetNomChamp(iChamp)));
          end;
        end;
        { Erreur ?}
        if (TobData.FieldExists('Error')) and (TobData.GetString('Error') <> '') then
        begin
          LastError    := 1;
          LastErrorMsg := TobData.GetString('Error');
        end;

        FreeAndNil(TobData);
      end
      else
      begin
        { Erreur ?}
        if (TobData.FieldExists('Error')) and (TobData.GetString('Error') <> '') then
        begin
//GP_20080208_TS_GP14792 >>>
          LastError    := 1;
          LastErrorMsg := TobData.GetString('Error');
//GP_20080208_TS_GP14792 <<<
        end;
      end;
    end;
  end;
{$ENDIF !PGIMAJVER}
end;
{$ENDIF !EAGLSERVER}
{$ENDIF GCGC}

{$IFNDEF EAGLSERVER}
procedure tTOMComm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    Ord('S'), Ord('s'): if Shift = [ssCtrl] then
                        begin
                          if Assigned(GetControl('MNLPDISPO')) and TmenuItem(GetControl('MNLPDISPO')).Enabled and TmenuItem(GetControl('MNLPDISPO')).Visible then
                            TmenuItem(GetControl('MNLPDISPO')).Click;
                          if Assigned(GetControl('MNDISPO')) and TmenuItem(GetControl('MNDISPO')).Enabled and TmenuItem(GetControl('MNDISPO')).Visible then
                            TmenuItem(GetControl('MNDISPO')).Click;
                        end;
    Ord('A'), Ord('a'): if Shift = [ssCtrl] then
                        begin
                          if Assigned(GetControl('MNLPARTICLE')) and TmenuItem(GetControl('MNLPARTICLE')).Enabled and TmenuItem(GetControl('MNLPARTICLE')).Visible then
                            TmenuItem(GetControl('MNLPARTICLE')).Click;
                        end;
  end;
  
  if Ecran is TFFiche then
    TFFiche(Ecran).FormKeyDown(Sender, Key, Shift)
  {$IFNDEF ERADIO}
    else if Ecran is TFSaisieList then
      TFSaisieList(Ecran).FormKeyDown(Sender, Key, Shift);
  {$ENDIF !ERADIO}
end;
{$ENDIF !EAGLSERVER}

function TTOMComm.EcranIs(const FormName: String): Boolean;
begin
  if Assigned(Ecran) then
    Result := Pos(UpperCase(Trim(FormName)), UpperCase(Ecran.Name)) = 1
  else
    Result := False
end;

{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 29/03/2007
Modifié le ... :   /  /    
Description .. : Retourne VRAI si on peut faire un ControlField
Mots clefs ... : 
*****************************************************************}
function tTOMComm.CanControl: Boolean;
begin
  Result := fDoControl = 0
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : permet de grouper les SetControlEnabled
Mots clefs ... :
*****************************************************************}
procedure tTOMComm.SetControlsEnabled(Const FieldsName: array of String; Const Enability: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
    SetControlEnabled(FieldsName[i], Enability);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/04/2003
Modifié le ... : 08/04/2003
Description .. : permet de grouper les SetControlVisible
Mots clefs ... :
*****************************************************************}
procedure tTOMComm.SetControlsVisible(Const FieldsName: array of String; Const Visibility: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
    SetControlVisible(FieldsName[i], Visibility);
end;

//GP_20080208_TS_GP14792 >>>
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/02/2008
Modifié le ... :   /  /
Description .. : Permet de gérer une erreur au niveau de la Tom si une
Suite ........ : exception est arrivée en cours de transaction.
Mots clefs ... :
*****************************************************************}
procedure tTOMComm.CheckIOError;
begin
  if (LastError = 0) and (V_Pgi.NbTransaction > 0) and (V_Pgi.ioError <> oeOk) then
  begin
    LastError := ErrTransactionTom;
    if Assigned(Ecran) then
      PgiError(TransactionErrorMessage())
    else if Assigned(FTob) then
      FTob.AddChampSupValeur('Error', TransactionErrorMessage(), CSTString)
  end
end;
//GP_20080208_TS_GP14792 <<<

Initialization
  registerclasses ( [ tTOMComm ] ) ;
end.
