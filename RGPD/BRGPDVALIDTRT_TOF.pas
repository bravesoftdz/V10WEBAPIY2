{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDVALIDTRT_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  , mul
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , FE_Main
  {$ENDIF EAGLCLIENT}
  , uTob
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  , CBPMcd
  , Htb97
  , uTOFComm
  , BRGPDUtils
  ;

function BLanceFiche_RGPDValidTrt(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDVALIDTRT = Class (tTOFComm)
  private
    Action       : T_RGPDActions;
    Population   : T_RGPDPopulation;
    PdfFile      : THEdit;
    GenerateDoc  : THEdit;
    Confirmation : THLabel;
    FileLabel    : THLabel;
    bValider     : TToolbarButton97;
    bInsert      : TToolbarButton97;
    sCode        : string;
    sLabel       : string;
    sLabel2nd    : string;
    CodeFileName : string;
    CurrentPath  : string;
    ResponseChoice : THRadioGroup;

    procedure PdfFile_OnChange(Sender : TObject);
    procedure PdfFile_OnElipsisClick(Sender : TObject);
    procedure bInsert_OnClick(Sender : TObject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;                                                                    

Implementation

uses
  wCommuns
  , UtilPGI
  , Vierge
  , BTPUtil
  , Paramsoc
  , FileCtrl
  , dialogs
  ;

function BLanceFiche_RGPDValidTrt(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  V_PGI.ZoomOle := True;
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
  V_PGI.ZoomOle := False;
end;

procedure TOF_BRGPDVALIDTRT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDVALIDTRT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDVALIDTRT.OnUpdate ;
var
  Cancel : boolean;
begin
  Inherited ;
  if PdfFile.Text = '' then
  begin
    Cancel := (Action = rgdpaConsentRequest);
    if not Cancel then
      Cancel := (PGIAsk('Vous n''avez pas sélectionné ' + iif(Action = rgdpaConsentRequest, 'modèle', 'demande') + '. Voulez-vous continuer ?', Ecran.Caption) <> mrYes);
  end else
  begin
    Cancel := (not FileExists(PdfFile.Text));
    if Cancel then
      PGIError(Format(TraduireMemoire('Le fichier %s n''existe pas.'), [PdfFile.Text]), Ecran.Caption);
  end;
  if (not Cancel) and (Action = rgdpaConsentResponse) then
  begin
    Cancel := (ResponseChoice.Value = 'A'); // A=Attente, V=Validée, R=Refusée
    if Cancel then
      PGIError(TraduireMemoire('Veuillez choisir une réponse.'), Ecran.Caption);
  end;
  if Cancel then
  begin
    TFVierge(Ecran).ModalResult := 0;
  end else
    TFVierge(Ecran).Retour := 'OK;'
                            + 'INPUT=' + PdfFile.Text
                            + iif(GenerateDoc.Text <> '', ';OUTPUT=' + GenerateDoc.Text, '')
                            + iif(Action = rgdpaConsentResponse, ';RESPONSE=' + ResponseChoice.Value, '');
                            ;
end ;

procedure TOF_BRGPDVALIDTRT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDVALIDTRT.OnArgument (S : String ) ;
var
  Who       : string;
  LDocGen   : THLabel;
  Cpt       : integer;
begin
  Inherited ;
  Action         := RGPDUtils.GetActionFromCode(GetArgumentString(S, 'ACTION'));
  Population     := RGPDUtils.GetPopulationFromCode(GetArgumentString(S, 'ORIGINE'));
  Confirmation   := THLabel(GetControl('CONFIRMATION'));
  FileLabel      := THLabel(GetControl('LABEL'));
  LDocGen        := THLabel(GetControl('LDOCGENERE'));
  PdfFile        := THEdit(GetControl('PDFFILE'));
  GenerateDoc    := THEdit(GetControl('DOCGENERE'));
  bValider       := TToolbarButton97(GetControl('BVALIDER'));
  bInsert        := TToolbarButton97(GetControl('BINSERT'));
  ResponseChoice := THRadioGroup(GetControl('RESPONSECONSENT'));
  sCode     := '';
  sLabel    := '';
  sLabel2nd := '';
  Who       := GetArgumentString(S, 'QUI', False);
  Cpt := 0;
  while Cpt < 3 do
  begin
    case Cpt of
      0 : sCode     := ReadTokenPipe(Who, '~');
      1 : sLabel    := ReadTokenPipe(Who, '~');
      2 : sLabel2nd := ReadTokenPipe(Who, '~');
    end;
    Inc(Cpt);
  end;
  if Action = rgdpaConsentRequest then
    sLabel := GetArgumentString(S, 'QTE', False);
  CodeFileName := GetArgumentString(S, 'CODEFILENAME', False);
  case Action of
    rgpdaDataExport
    , rgpdaAnonymization
    , rgdpaDataRectification :  begin
                                  FileLabel.Caption := TraduireMemoire('Sélection de la demande');
                                  //PdfFile.DataType  := 'OPENFILE(*.PDF;*.*)';
                                  Ecran.Caption     := Format('%s - %s %s %s', [  RGPDUtils.GetLabelFromAction(Action)
                                                                                , RGPDUtils.GetLabelFromPopulation(Population)
                                                                                , sLabel
                                                                                , sLabel2nd
                                                                               ]);
                                end;
    rgdpaConsentRequest
    , rgdpaConsentResponse   :  begin
                                  FileLabel.Caption := iif(Action = rgdpaConsentRequest, TraduireMemoire('Sélection du modèle'), TraduireMemoire('Sélection du document de réponse'));
                                  //PdfFile.DataType  := 'OPENFILE(*.DOTX)';
                                  Ecran.Caption     := Format(TraduireMemoire('%s pour %s %s'), [  RGPDUtils.GetLabelFromAction(Action)
                                                                                                 , sLabel
                                                                                                 , RGPDUtils.GetLabelFromPopulationM(Population)
                                                                                                ]);
                                end;
  end;
  LDocGen.Visible        := ((Action = rgdpaConsentRequest) or (Action = rgdpaConsentResponse));
  LDocGen.Caption        := iif(Action = rgdpaConsentRequest, 'Répertoire des documents à générer', 'Réponse');
  GenerateDoc.Visible    := (Action = rgdpaConsentRequest);
  bInsert.Visible        := (Action = rgdpaConsentRequest);
  ResponseChoice.Visible := (Action = rgdpaConsentResponse);
  GenerateDoc.Text       := iif(Action = rgdpaConsentRequest, RGPDUtils.GetRgpdPSoc(rgpdtGenerationConsent), '');
  bValider.Enabled       := (Action <> rgdpaConsentRequest);
  bInsert.OnClick        := bInsert_OnClick;
  PdfFile.OnChange       := PdfFile_OnChange;
  PdfFile.OnElipsisClick := PdfFile_OnElipsisClick;
  TFVierge(Ecran).Retour := '';
end ;

procedure TOF_BRGPDVALIDTRT.OnClose ;
begin
  Inherited ;
  SetCurrentDir(CurrentPath);
end ;

procedure TOF_BRGPDVALIDTRT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDVALIDTRT.OnCancel () ;
begin
  Inherited ;
  TFMul(Ecran).Retour := 'CANCEL';
end ;

procedure TOF_BRGPDVALIDTRT.PdfFile_OnChange(Sender: TObject);
begin
  bValider.Enabled := ((Action <> rgdpaConsentRequest) or ((Action = rgdpaConsentRequest) and (PdfFile.Text <> '')));
end;

procedure TOF_BRGPDVALIDTRT.PdfFile_OnElipsisClick;
var
  GetPath : TOpenDialog;
begin
  GetPath := TOpenDialog.Create(Ecran);
  try
    if Action = rgdpaConsentRequest then
    begin
      GetPath.FileName   := RGPDUtils.GetRgpdPSoc(rgpdtTemplateConsent) + '*.DOTX';
      GetPath.Title      := 'Choix du modèle';
      GetPath.Filter     := '*.DOTX';
    end else
    begin
      GetPath.FileName   := RGPDUtils.GetRgpdPSoc(rgpdtIncomingRequest) + '*.PDF';
      GetPath.Title      := 'Choix de la demande';
      GetPath.Filter     := '*.PDF';
    end;
    GetPath.Options    := GetPath.Options + [ofNoChangeDir];
    if GetPath.Execute then
      PdfFile.Text := GetPath.FileName;
  finally
    FreeAndNil(GetPath);
  end;
end;

procedure TOF_BRGPDVALIDTRT.bInsert_OnClick(Sender: TObject);
var
  InputFilePath  : string;
  CodeFileNameL  : string;
  TableName      : string;
  FieldsList     : string;
  FieldsListSql  : string;
  FieldsList1    : string;
  FieldsListSql1 : string;
  Where          : string;
  TobFields      : TOB;

  procedure AddAdress(sTableName : string);
  begin
    Publipost.SetFieldsList(Publipost.GetPrefixFromTableName(TableName), sTableName, FieldsList1, FieldsListSql1);
    Publipost.AddFieldsInTob(TobFields, sTableName, FieldsListSql1, Where, True);
    FieldsList := FieldsList + ';' + FieldsList1;
  end;

begin
  TobFields := TOB.Create('', nil, -1);
  try
    InputFilePath := IncludeTrailingPathDelimiter(RGPDUtils.GetRgpdPSoc(rgpdtTemplateConsent)) + 'NewDoc.doc';
    CodeFileNameL := CodeFileName;
    TableName     := RGPDUtils.GetTableNameFromPopulation(Population);
    while CodeFileNameL <> '' do
      Where := Where + ' AND ' + ReadTokenPipe(CodeFileNameL, '|') + ' = "' + ReadTokenPipe(sCode, '|') + '"';
    Where := Copy(Where, 5, length(Where));
    Publipost.SetFieldsList(Publipost.GetPrefixFromTableName(TableName), TableName, FieldsList, FieldsListSql);
    Publipost.AddFieldsInTob(TobFields, TableName, FieldsListSql, Where, True);
    { Pour les contacts, il faut trouver la 1ère adresse éventuellement associée }
    if (Population = rgpdpContact) then
    begin
      Where := StringReplace(Where, 'C_TIERS', 'ADR_REFCODE', [rfReplaceAll]);
      Where := StringReplace(Where, 'C_NUMEROCONTACT', 'ADR_NUMEROCONTACT', [rfReplaceAll]);
      AddAdress('ADRESSES');
    end else
    { Pour les utilisateurs, il faut trouver l'adresse associée au salarié si existe }
    if (Population = rgpdpUser) then 
    begin
      Where := ' PSA_SALARIE = ' + Copy(Where, Pos('US_AUXILIAIRE', Where) + 16, Length(Where));
      AddAdress('SALARIES');
    end;
    Publipost.NewModel(InputFilePath, TobFields, FieldsList);
  finally
    FreeAndNil(TobFields);
  end;
end;

Initialization
  registerclasses ( [ TOF_BRGPDVALIDTRT ] ) ;
end.


