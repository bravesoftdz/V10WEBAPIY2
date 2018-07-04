unit UtofRTSaisieInfosEsker;

interface
uses  Controls,Classes,forms,sysutils,HMsgBox,Hstatus,M3FP,
      HCtrls,HEnt1,UTOF,ParamSoc,UtilRT,UTob,EntRT,HQry,HTB97,Vierge,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul
{$ELSE}
      HDB,db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,Mul
{$ENDIF}
      ,EskerInterface
      ;

Type
    TOF_RTSaisieInfosEsker = Class (TOF)
     private
        TobInfos,Tobadr : Tob;
        ErreurSaisie : integer;
        stDocument : string;
        procedure BPREVISU_OnClick(Sender: TObject);
        procedure MajTobInfos;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     END ;

Function RTLanceFiche_RTSaisieInfosEsker(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}        'Vous devez saisir un libellé'
          {2}        ,'La date ne peut pas être antérieure à la date du jour'
          {3}        ,'Vous devez saisir un objet'
          );

implementation

uses SubmissionService,UtilPGI;

procedure TOF_RTSaisieInfosEsker.OnArgument(Arguments : String ) ;
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
inherited ;
//Tobinfos := LaTob;
ErreurSaisie:=0;
Repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere<>'' then
    begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='DOCUMENT' then
              stDocument := ValMul;
           if ChampMul='TOBADR' then
              TobAdr := Tob(Valeuri(ValMul));
           if ChampMul='TOBINFOS' then
              Tobinfos := Tob(Valeuri(ValMul));
        end
    end;
until  Critere='';
if Assigned(GetControl('BPREVISU')) then
   TToolbarButton97(GetControl('BPREVISU')).OnClick := BPREVISU_OnClick;

TobInfos.AddChampSupValeur ('VALIDATION','N',False);
if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
  begin
  TobInfos.AddChampSupValeur ('NOMENVOI','',False);
  TobInfos.AddChampSupValeur ('AFFRANCHISSEMENT','',False);
  TobInfos.AddChampSupValeur ('TAILLEENV','',False);
  TobInfos.AddChampSupValeur ('PORTEADRESSE','',False);
  TobInfos.AddChampSupValeur ('RECTOVERSO','',False);
  TobInfos.AddChampSupValeur ('COULEUR','',False);
  TobInfos.AddChampSupValeur ('DATEENVOI',V_PGI.DateEntree,False);
  TobInfos.AddChampSupValeur ('NOMEMETTEUR',VH_RT.RTNomResponsable,False);
  TobInfos.AddChampSupValeur ('NOMSOCIETE',GetParamSocSecur('SO_LIBELLE',''),False);
  if TobInfos.GetString ('TYPECOURRIER') = 'C' then THValComboBox(GetControl('AFFRANCHISSEMENT')).Plus := 'AND (CO_CODE = "001" OR CO_CODE = "002")';
  end;
if TFVierge(Ecran).name = 'RTSAISIEEMAIL' then
  begin
  TobInfos.AddChampSupValeur ('OBJET','',False);
  TobInfos.AddChampSupValeur ('TEXTESAISI','',False);
  TobInfos.AddChampSupValeur ('IMPORTANCE','',False);
  TobInfos.AddChampSupValeur ('DATEENVOI',V_PGI.DateEntree,False);
  TobInfos.AddChampSupValeur ('NOMEMETTEUR',VH_RT.RTNomResponsable,False);
  TobInfos.AddChampSupValeur ('NOMSOCIETE',GetParamSocSecur('SO_LIBELLE',''),False);
  end;
if TFVierge(Ecran).name = 'RTSAISIEFAX' then
  begin
  TobInfos.AddChampSupValeur ('OBJET','',False);
  TobInfos.AddChampSupValeur ('TEXTESAISI','',False);
  TobInfos.AddChampSupValeur ('DATEENVOI',V_PGI.DateEntree,False);
  TobInfos.AddChampSupValeur ('NOMEMETTEUR',VH_RT.RTNomResponsable,False);
  TobInfos.AddChampSupValeur ('NOMSOCIETE',GetParamSocSecur('SO_LIBELLE',''),False);
  TobInfos.AddChampSupValeur ('FAXSOCIETE',CleTelephone(GetParamSocSecur('SO_FAX','')),False);
  end;
if TFVierge(Ecran).name = 'RTSAISIESMS' then
  begin
  TobInfos.AddChampSupValeur ('TEXTESAISI','',False);
  end;
end;

procedure TOF_RTSaisieInfosEsker.OnLoad;
begin
inherited;
if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
  begin
  SetControlText('AFFRANCHISSEMENT', '001');
  SetControlText('TAILLEENV', '002');
  end;
end;

procedure TOF_RTSaisieInfosEsker.OnUpdate;
begin
inherited;
if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
  begin
  if GetControlText('NOMENVOI') = '' then
    begin
    Lasterror:=1; //ErreurSaisie:=1;
    LastErrorMsg:=TexteMessage[LastError];
    SetFocusControl('NOMENVOI') ;
    TFVierge(Ecran).ModalResult := 0;
    exit;
    end;
  end;
if (TFVierge(Ecran).name = 'RTSAISIEEMAIL') or (TFVierge(Ecran).name = 'RTSAISIEFAX') then
  begin
  if GetControlText('OBJET') = '' then
    begin
    Lasterror:=3; //ErreurSaisie:=1;
    LastErrorMsg:=TexteMessage[LastError];
    SetFocusControl('OBJET') ;
    TFVierge(Ecran).ModalResult := 0;
    exit;
    end;
  end;
if TFVierge(Ecran).name = 'RTSAISIESMS' then
  begin
  if GetControlText('TEXTESAISI') = '' then
    begin
    Lasterror:=1; //ErreurSaisie:=1;
    LastErrorMsg:=TexteMessage[LastError];
    SetFocusControl('TEXTESAISI') ;
    TFVierge(Ecran).ModalResult := 0;
    exit;
    end;
  end;
if TFVierge(Ecran).name <> 'RTSAISIESMS' then
  begin
  if StrToDate(GetControlText ('DATEENVOI')) < V_PGI.DateEntree then
    begin
    Lasterror:=2; //ErreurSaisie:=1;
    LastErrorMsg:=TexteMessage[LastError];
    SetFocusControl('DATEENVOI') ;
    TFVierge(Ecran).ModalResult := 0;
    exit;
    end;
  end;
MajTobInfos;
if (TobInfos.FieldExists ('TYPECOURRIER')) and (TobInfos.GetString ('TYPECOURRIER') = 'M') then
  TobInfos.SetString ('VALIDATION','Y')
else
  TobInfos.SetString ('VALIDATION',GetParamSocSecur('SO_VALIDATIONENVOI','Y'));
TFVierge(Ecran).Retour := 'SAISIE';
//LaTob := TobInfos;
end;

procedure TOF_RTSaisieInfosEsker.OnClose;
begin
inherited;
end;

procedure TOF_RTSaisieInfosEsker.MajTobInfos ;
begin
if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
  begin
  TobInfos.SetString ('NOMENVOI',GetControlText('NOMENVOI'));
  TobInfos.SetString ('AFFRANCHISSEMENT',GetControlText('AFFRANCHISSEMENT'));
  TobInfos.SetString ('TAILLEENV',GetControlText('TAILLEENV'));
  TobInfos.SetString ('PORTEADRESSE',GetControlText('PORTEADRESSE'));
  TobInfos.SetString ('RECTOVERSO',GetControlText('RECTOVERSO'));
  TobInfos.SetString ('COULEUR',GetControlText('COULEUR'));
  TobInfos.SetString ('DATEENVOI',GetControlText('DATEENVOI'));
  end;
if TFVierge(Ecran).name = 'RTSAISIEEMAIL' then
  begin
  TobInfos.SetString ('OBJET',GetControlText('OBJET'));
  TobInfos.SetString ('TEXTESAISI',GetControlText('TEXTESAISI'));
  TobInfos.SetString ('IMPORTANCE',GetControlText('IMPORTANCE'));
  TobInfos.SetString ('DATEENVOI',GetControlText('DATEENVOI'));
  end;
if TFVierge(Ecran).name = 'RTSAISIEFAX' then
  begin
  TobInfos.SetString ('OBJET',GetControlText('OBJET'));
  TobInfos.SetString ('TEXTESAISI',GetControlText('TEXTESAISI'));
  TobInfos.SetString ('DATEENVOI',GetControlText('DATEENVOI'));
  end;
if TFVierge(Ecran).name = 'RTSAISIESMS' then
  begin
  TobInfos.SetString ('TEXTESAISI',GetControlText('TEXTESAISI'));
  end;
end;

procedure TOF_RTSaisieInfosEsker.BPREVISU_OnClick(Sender: TObject) ;
var Identifiant : SubmissionResult;
begin
if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
  begin
  if GetControlText('NOMENVOI') = '' then exit;
  end;
if TFVierge(Ecran).name = 'RTSAISIEFAX' then
  begin
  if GetControlText('OBJET') = '' then exit;
  end;
MajTobInfos;
TobInfos.SetString ('VALIDATION','Y');
if stDocument <> '' then
  begin
  if TFVierge(Ecran).name = 'RTSAISIECOURRIER' then
    Identifiant := SendEsker (0,TobInfos,stDocument,TobAdr)
  else
    Identifiant := SendEsker (1,TobInfos,stDocument,TobAdr);
  if Identifiant <> Nil then QueryEsker (Identifiant);
  end;
end;

Function RTLanceFiche_RTSaisieInfosEsker(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result := AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

Initialization
registerclasses([TOF_RTSaisieInfosEsker]);
end.
