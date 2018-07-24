unit FormsName;

Interface

uses
  Hent1
  , BRGPDUtils
  ;

const
  frm_RGPDRepository      = 'BRGPDREFERENTIEL';
  frm_RGPDThirdMul        = 'BRGPDTIERSMUL';
  frm_RGPDTrtValid        = 'BRGPDVALIDTRT';
  frm_RGPDResourceMul     = 'BRGPDRESSOURCEMUL';
  frm_RGPDUtilisatMul     = 'BRGPDUTILISATMUL';
  frm_RGPDSuspectMul      = 'BRGPDSUSPECTMUL';
  frm_RGPDContactMul      = 'BRGPDCONTACTMUL';
  frm_RGPDSensibilization = 'BRGPDSENSIBILISAT';
  frm_ThirdCliPro         = 'GCTIERS';
  frm_Resource            = 'BTRESSOURCE';
  frm_Utilisat            = 'YYUTILISAT';
  frm_Suspect             = 'RTSUSPECTS';
  frm_JnalEvent           = 'YYJNALEVENT';
  frm_Contact             = 'YYCONTACT';
  frm_WSAllowedTable      = 'BTWSTABLEAUTO';

type
  OpenForm = class
    private
      class function SetArgument(Argument : string) : string;

    public
      class function SetWindowCaption(ParentNumber, TagNumber : integer; Action : T_RGPDActions=rgpdaNone; Population : T_RGPDPopulation=rgpdpNone) : string;
      class function CliPro(Auxiliary, ThirdType : string; Argument : string=''; ForceConsult : Boolean=False) : string;
      class function Resource(ResourceCode, ResourceType : string; Argument : string=''; ForceConsult : Boolean=False) : string;
      class function User(UserCode : string; ForceConsult : Boolean=False) : string;
      class function Suspect(SuspectCode : string; ForceConsult : Boolean=False) : string;
      class function Contact(ContactType, AuxiliaryCode : string; ContactNumber : integer; Argument : string=''; ForceConsult : Boolean=False) : string;
      class function JnalEvent(Params : string='') : string;
      class function RGPDSensibilisation : string;
      class function RGPDWindows(Population : T_RGPDPopulation; TagMother, TagNumber : integer; Action : T_RGPDActions) : string;
      class function WSCreationAllowedTable : string;
    end;

implementation

uses
  UFonctionsCBP
  , BTPUtil
  , Fe_Main
  , uTOFComm
  , ConfidentAffaire
  , HCtrls
  , SysUtils
  , wCommuns
  , uDbxDataSet
  , BRGPDTIERSMUL_TOF
  , BRGPDRESSOURCEMUL_TOF
  , BRGPDUTILISATMUL_TOF
  , BRGPDSUSPECTMUL_TOF
  , BRGPDCONTACTMUL_TOF
  , BTWSTABLEAUTO_TOF
  , CommonTools
  ;

class function OpenForm.SetArgument(Argument: string): string;
begin
  if (Argument <> '') and (Copy(Argument, 1, 1) <> ';') then
    Result := ';' + Argument
  else
    Result := Argument;
end;

class function OpenForm.SetWindowCaption(ParentNumber, TagNumber : integer; Action : T_RGPDActions=rgpdaNone; Population : T_RGPDPopulation=rgpdpNone) : string;
var
  Sql : string;
  Qry : TQuery;
  IsConsent : boolean;
  Num       : Integer;
begin
  IsConsent := ((Action = rgdpaConsentRequest) or (Action = rgdpaConsentResponse));
  Num := 0;
  Sql := 'SELECT MN_LIBELLE FROM MENU WHERE MN_TAG IN(' + IntToStr(ParentNumber) + ', ' + IntToStr(TagNumber) + ')';
  Qry := OpenSQL(Sql, True);
  try
    while not Qry.Eof do
    begin
      Inc(Num);
      Result := Result + ' - ' + Qry.Fields[0].AsString;
      if (IsConsent) and (Num > 1) then
        Result := Result + ' ' + RGPDUtils.GetLabelFromPopulation(Population);
      Qry.Next;
    end;
  finally
    Ferme(Qry);
  end;
  Result := Copy(Result, 4, Length(Result));
  Result := 'WINDOWCAPTION=' + Result;
end;

class function OpenForm.CliPro(Auxiliary, ThirdType : string; Argument : string=''; ForceConsult : Boolean=False) : string;
var
  stAction   : string;
  stArgument : string;
begin
  if (Auxiliary <> '') and (ThirdType <> '') then
  begin
    stAction   := 'ACTION=MODIFICATION';
    stArgument := Argument;
    if ThirdType = 'PRO' then
    begin
      if not ExJaiLeDroitConcept(TConcept(bt510),False) then
        stAction:= 'ACTION=CONSULTATION';
    end else
      if not ExJaiLeDroitConcept(TConcept(bt511),False) then
        stAction:= 'ACTION=CONSULTATION';
    stArgument := SetArgument(Argument);
    if (ForceConsult) or (not Tools.CanInsertedInTable('TIERS'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV})) then
      stAction:= 'ACTION=CONSULTATION';
    Result := AGLLanceFiche('GC', frm_ThirdCliPro, '', Auxiliary, stAction + ';MONOFICHE;T_NATUREAUXI=' + ThirdType + stArgument);
  end else
    Result := '';
end;

class function OpenForm.Resource(ResourceCode, ResourceType : string; Argument : string=''; ForceConsult : Boolean=False) : string;
var
  stAction   : string;
  stArgument : string;
begin
  if ResourceCode <> '' then
  begin
    stAction   := ';ACTION=MODIFICATION';
    stArgument := Argument;
    if AGLJaiLeDroitFiche(['RESSOURCE', stAction], 2) then
    begin
      if ForceConsult then
        stAction:= ';ACTION=CONSULTATION';
      stArgument := SetArgument(Argument);
      Result := AGLLanceFiche('BTP', frm_Resource, '', ResourceCode, 'TYPERESSOURCE=' + ResourceType + stAction + stArgument);
    end else
      Result := '';
  end else
    Result := '';
end;

class function OpenForm.User(UserCode: string; ForceConsult : Boolean=False): string;
begin
  if (UserCode <> '') and (JaiLeDroitTag(60202)) then
    Result := AGLLanceFiche('YY', frm_Utilisat, '', UserCode, iif(not ForceConsult, 'ACTION=MODIFICATION', 'ACTION=CONSULTATION'));
end;

class function OpenForm.Suspect(SuspectCode: string; ForceConsult : Boolean=False): string;
begin
  if (SuspectCode <> '') and (JaiLeDroitTag(92106)) then
    Result := AGLLanceFiche('RT', frm_Suspect, '', SuspectCode, 'MONOFICHE;' + iif(not ForceConsult, 'ACTION=MODIFICATION', 'ACTION=CONSULTATION'));
end;

class function OpenForm.Contact(ContactType, AuxiliaryCode : string; ContactNumber : integer; Argument : string=''; ForceConsult : Boolean=False) : string;
var
  stArgument : string;
begin
  if (ContactType <> '') and (AuxiliaryCode <> '') and (ContactNumber > 0) and (JaiLeDroitTag(92106)) then
  begin
    stArgument := SetArgument(Argument);
    Result     := AGLLanceFiche('YY', frm_Contact, ContactType + ';' + AuxiliaryCode, IntToStr(ContactNumber), iif(not ForceConsult, 'ACTION=MODIFICATION', 'ACTION=CONSULTATION') + stArgument) ;
  end;
end;

class function OpenForm.JnalEvent(Params : string): string;
begin
  Result := AGLLanceFiche('YY', frm_JnalEvent, '', '', Params);
end;

class function OpenForm.RGPDSensibilisation : string;
begin
  V_PGI.ZoomOle := True;
  Result := AGLLanceFiche('BTP', frm_RGPDSensibilization, '', '', '');
  V_PGI.ZoomOle := False;

end;

class function OpenForm.RGPDWindows(Population: T_RGPDPopulation; TagMother, TagNumber: integer; Action: T_RGPDActions): string;
begin
  case Population of
    rgpdpThird    : BLanceFiche_RGPDThirdMul   ('BTP', frm_RGPDThirdMul   , '', '', OpenForm.SetWindowCaption(TagMother, TagNumber, Action, Population)+ ';ACTION=' + RGPDUtils.GetCodeFromAction(Action));
    rgpdpResource : BLanceFiche_RGPDResourceMul('BTP', frm_RGPDResourceMul, '', '', OpenForm.SetWindowCaption(TagMother, TagNumber, Action, Population)+ ';ACTION=' + RGPDUtils.GetCodeFromAction(Action));
    rgpdpUser     : BLanceFiche_RGPDUtilisatMul('BTP', frm_RGPDUtilisatMul, '', '', OpenForm.SetWindowCaption(TagMother, TagNumber, Action, Population)+ ';ACTION=' + RGPDUtils.GetCodeFromAction(Action));
    rgpdpSuspect  : BLanceFiche_RGPDSuspectMul ('BTP', frm_RGPDSuspectMul , '', '', OpenForm.SetWindowCaption(TagMother, TagNumber, Action, Population)+ ';ACTION=' + RGPDUtils.GetCodeFromAction(Action));
    rgpdpContact  : BLanceFiche_RGPDContactMul ('BTP', frm_RGPDContactMul , '', '', OpenForm.SetWindowCaption(TagMother, TagNumber, Action, Population)+ ';ACTION=' + RGPDUtils.GetCodeFromAction(Action));
  end;
end;

class function OpenForm.WSCreationAllowedTable: string;
begin
  Result := BLanceFiche_WSTablesAutorisees('BTP', frm_WSAllowedTable, '', '', SetWindowCaption(-14811, 148906));
end;

end.
