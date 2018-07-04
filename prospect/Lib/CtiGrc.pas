unit Ctigrc;

interface

uses
  SysUtils, Classes, Controls, Forms, 
  StdCtrls, OleCtrls, UtilCti,hctrls,comctrls,
  HTB97,ParamSoc
{$IFDEF EAGLCLIENT}
      , Maineagl
{$ELSE}
      , FE_main
{$ENDIF}
,CCACONTROLSLib_TLB,HEnt1, Mask ;

Const
  Avance : integer = 1;
  Simple : integer = 0;
  ModeAppelSortant : integer = 1 ;
  ModeAppelEntrant : integer = 2 ;

type
  TResultatAppel =(raEnCours,raReponse,raOccupe,raRefus,raAbort,raFinCommunication);

  TCCaForm = class(TForm)
    LOG: TMemo;
    APPELENTRANT: TGroupBox;
    NOTELAPPELANT: TEdit;
    LNOTELAPPELANT: TLabel;
    LTIERSAPPELANT: TLabel;
    TIERSAPPELANT: TEdit;
    LIBELLEAPPELANT: TEdit;
    LLIBELLEAPPELANT: TLabel;
    APPELSORTANT: TGroupBox;
    LNOTELAPPELE: TLabel;
    NOTELAPPELE: TEdit;
    LTIERSAPPELE: TLabel;
    LIBELLEAPPELE: TEdit;
    LLIBELLEAPPELE: TLabel;
    TIERSAPPELE: THCritMaskEdit;
    Dock972: TDock97;
    PBouton: TToolWindow97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    NATUREAPPELE: TEdit;
    NATUREAPPELANT: TEdit;
    AUXIAPPELANT: TEdit;
    AUXIAPPELE: TEdit;
    RichEdit1: TRichEdit;
    BCONTACT: TToolbarButton97;
    BTIERS: TToolbarButton97;
    EFFACER: TToolbarButton97;
    BACTIONS: TToolbarButton97;
    Label2: TLabel;
    LNOMAPPELANT: TLabel;
    NOMAPPELE: TEdit;
    LNOMAPPELE: TLabel;
    NUMCONTACTAPPELE: TEdit;
    NUMCONTACTAPPELANT: TEdit;
    //AxApplication1: TAxApplication;
    //AxLine1: TAxLine;
    NOMAPPELANT: TEdit;
    AxApplication1: TAxApplication;
    AxLine1: TAxLine;
    procedure FormShow(Sender: TObject);
    procedure FermerClick(Sender: TObject);
    procedure AppelClick(Sender: TObject);
    procedure AttenteClick(Sender: TObject);
    procedure RepriseClick(Sender: TObject);
    procedure AppelListeTiers(Sender: TObject);
    procedure DecrocherClick(Sender: TObject);
    procedure RaccrocherClick(Sender: TObject);
    procedure EffacerClick(Sender: TObject);

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    Function  RTCtiSelectionClient ( appele : WideString; AffMul : Boolean) : String;
    Function  RTCtiDetailFicheClient ( StTiers , StNat : String ) : String;
    Function  RTCtiReprendFicheClient ( TypeMontee, StTiers , StNat : String ) : String;
    Procedure RTCTINumeroteClient( Sender: TObject ; TelephoneAppele,StTiers,StNom,StNat,StAuxi : String );
    procedure AddLog (StLog : String);
    procedure AlimChamps( Appelant : Boolean; RetCti : String );
    procedure ContactClick(Sender: TObject);
    procedure TiersClick(Sender: TObject);
    procedure ActionsClick(Sender: TObject);
    procedure AfficheCoordonnees( Appelant : Boolean; Tiers,Libelle,Nature,Auxi,Num,Nom : String );
    { Déclarations privées }
    procedure OnLineDeleteCall(Sender: TObject; callId : integer );
    procedure OnLineChangeCall(Sender: TObject; const pCall: IDataOfCall);
    //procedure OnFailed(Sender: TObject);
    //procedure AppelEntrant(Sender: TObject; const pCall: IDataOfCall; callType: AxCallType);
    procedure OnLineStateChange (Sender: TObject; State: Integer);
  public
    { Déclarations publiques }
    NumAppel : integer;
    NoTel : String;
    AppelEnCours : Boolean ; // true si appel en cours ou fiche client affichée, false sinon
    FicheMontee : Boolean ; // pour gérer le raccrochage client
    Correlateur : String; // correlator data
  end;

Procedure RTLanceCCA ;

var
    Fo_CCA : TCCaForm;
    CtiErreurConnexion: Boolean;
    CtiResultatAppel : Smallint; // etat appel sortant
    CtiRaccrocherFiche : Boolean ; // true si l'on a raccroché à partir de la fiche client ou le correspondant a raccroché
    CCAHeureDeb : TDateTime;
    CCAModeAppel : Integer ; // 1 = Appel sortant; 2 = Appel Entrant
    CtiNumAct : Integer ; // No action crée
    RTFormCti : TForm;
implementation

{$R *.DFM}
Procedure RTLanceCCA ;
Begin
     Fo_CCA := TCCaForm.Create (Application);
     {Try
         Fo_CCA.Show;
     Finally
         Fo_CCA.free;}
     //Application.Minimize;
     Fo_CCA.ShowModal;

     //End;
end;

procedure TCCaForm.FormShow(Sender: TObject);
begin
AxApplication1.Connect;
AxApplication1.CCAgentVisibility:=AxALC_CALLSBAR_VISIBLE or AxALC_AGENTBAR_VISIBLE or AxALC_TELEPHONEBAR_VISIBLE;
AxLine1.Connect;
{ modif fiche pour CCA }
{$IFNDEF CCATEST}
APPELSORTANT.Visible:=false;
APPELENTRANT.Visible:=false;
LOG.Align:=alClient;
{$ENDIF}

Correlateur:='';
NoTel:='';
AppelEnCours:=False ;
//CtiRaccrocherFiche:=False ;
//FicheMontee:=false ;
////AfficheFeu;
//NumAppel:=AgentPhone1.AgentServiceState;
{$IFDEF CCATEST}
NumAppel:=1;
{$ELSE}
NumAppel:=0;
{$ENDIF}
CtiResultatAppel:=4;
CtiErreurConnexion:=False;
CCAHeureDeb:=0;
RTFormCti:=Nil;
end;

Procedure TCCaForm.RTCTINumeroteClient( Sender: TObject ; TelephoneAppele,StTiers,StNom,StNat,StAuxi : String );
begin
//AppelEnCours:=True ;
////AfficheFeu;
CCAModeAppel := ModeAppelSortant;
Notel:=TelephoneAppele;
if Notel <> '' then
   begin
   TIERSAPPELE.Text:= StTiers;
   LIBELLEAPPELE.Text:= StNom;
   NATUREAPPELE.Text:= StNat;
   AUXIAPPELE.Text:=StAuxi;
   end;
AppelClick (Sender);
end;

procedure TCCaForm.FermerClick(Sender: TObject);
begin
Close;
end;

procedure TCCaForm.OnLineStateChange (Sender: TObject; State: Integer);
begin
AddLog('--> State= '+IntToStr(State));
if state = 2 then
  begin
       AddLog(TraduireMemoire('--> Agent prêt '));
       AddLog(' ') ;
  end;
end;

procedure TCCaForm.AppelClick(Sender: TObject);
begin
NOTELAPPELE.text:=NoTel;
if NOTELAPPELE.text<>'' then
    begin
    if GetParamsocSecur('SO_RTCTILIGNEEXT','00') <> '' then
       NOTELAPPELE.text:=GetParamsocSecur('SO_RTCTILIGNEEXT','00')+NOTELAPPELE.text;
    AddLog(TraduireMemoire('--> Appel no ')+NOTELAPPELE.text);
    AddLog(' ') ;
    AxLine1.MakeCall(NOTELAPPELE.text,Correlateur) ;
    end;
end;

procedure TCCaForm.AppelListeTiers(Sender: TObject);
begin
//NOTELAPPELE.text:=DispatchRecherche(TIERSAPPELE,2,'','','');
RTCtiSelectionClient ( TIERSAPPELE.text, True);
end;

procedure TCCaForm.DecrocherClick(Sender: TObject);
begin
{$IFDEF CCATEST}
if NumAppel <> 0 then
{$ELSE}
if AxLine1.IsTakeCallPossible[AxLine1.ActiveCall] <> 0 then
{$ENDIF}
    begin
    //AppelEnCours:=True ;
    ////AfficheFeu;
    CCAModeAppel:=ModeAppelEntrant;
    AxLine1.TakeCall(AxLine1.ActiveCall);
    end;
end;

procedure TCCaForm.RaccrocherClick(Sender: TObject);
begin
{$IFDEF CCATEST}
if NumAppel <> 0 then
{$ELSE}
if AxLine1.IsReleaseCallPossible[AxLine1.ActiveCall] <> 0 then
{$ENDIF}
    begin
    AxLine1.ReleaseCall(AxLine1.ActiveCall);
    NumAppel:=0;
    if AppelEnCours then
        CtiRaccrocherFiche := True
    else
        CtiRaccrocherFiche := False ;
    //AfficheFeu;
    end;
end;

procedure TCCaForm.AttenteClick(Sender: TObject);
begin
{$IFDEF CCATEST}
if NumAppel <> 0 then
{$ELSE}
if AxLine1.IsHoldCallPossible <> 0 then
{$ENDIF}
    begin
    AxLine1.HoldCall;
    end;
end;

procedure TCCaForm.RepriseClick(Sender: TObject);
begin
{$IFDEF CCATEST}
if NumAppel <> 0 then
{$ELSE}
if AxLine1.IsReleaseCallPossible[AxLine1.ActiveCall] <> 0 then
{$ENDIF}
    begin
    AxLine1.ReleaseCall(AxLine1.ActiveCall);
//if AxLine1.IsRetrieveCallPossible <> 0 then
    //AxLine1.RetrieveCall;
    end;
end;


procedure TCCaForm.OnLineDeleteCall(Sender: TObject; callId : integer);
var HeureFin : TDateTime;
begin
AddLog(TraduireMemoire('--> Fin communication '));
// génération d'une action
HeureFin:=Time;
if (CCAModeAppel = ModeAppelEntrant ) and (TIERSAPPELANT.Text <> '') then
   RTCTIGenereAction(False,TIERSAPPELANT.Text,AUXIAPPELANT.Text,CCAHeureDeb,HeureFin,'',CCAModeAppel,StrToInt(NUMCONTACTAPPELANT.Text));
end;

procedure TCCaForm.OnLineChangeCall(Sender: TObject; const pCall: IDataOfCall);
var appelant,RetCti : String;
begin
{$IFNDEF CCATEST}
  if ( pCall.Incoming <> 0) and (pCall.ACDCall <> 0) and (pCall.State = AxALC_CALL_STATE_TALKING) then
  begin
  NOTELAPPELANT.text:= pCall.Number;
{$ENDIF}
  appelant:=NOTELAPPELANT.text;
  {if CCAModeAppel = ModeAppelEntrant then
     NOTELAPPELANT.text:=appelant;}
  CCAModeAppel := ModeAppelEntrant;
  //AddLog('--> Demande d''identification Appel '+IntToStr(RefAppel));
  AddLog(TraduireMemoire('--> Appel entrant'));
  AddLog(TraduireMemoire('    Appelant : ')+NOTELAPPELANT.text);
  if CCAModeAppel = ModeAppelEntrant then
      begin
      //AppelEnCours:=True ;
      ////AfficheFeu;
      FicheMontee:=true;
      CtiRaccrocherFiche:=False ;
      TIERSAPPELANT.Text:='';
      NATUREAPPELANT.Text:='';
      LIBELLEAPPELANT.Text:= '';
      CCAHeureDeb:=Time;
      RetCti:=RTCtiAfficheClient ( appelant,'appel;COUPLAGE;CCA');
      FicheMontee:=False;
      if (CtiRaccrocherFiche) {and (CtiResultatAppel <> 0)} then
          begin
          //AppelEnCours:=False ;
          ////AfficheFeu;
          //CtiRaccrocherFiche:=False ;
          end;

      if RetCti <> ';;;' then
         AlimChamps(True,RetCti)
      else
          // pas de fiche tiers, on recherche une fiche contact
          begin
          RetCti:=RTCtiAfficheContact ( appelant,'appel;COUPLAGE;CCA;CTI');
          FicheMontee:=False;
          if (CtiRaccrocherFiche) {and (CtiResultatAppel <> 0)} then
              begin
              //AppelEnCours:=False ;
              ////AfficheFeu;
              //CtiRaccrocherFiche:=False ;
              end;

          if RetCti <> ';;;0;' then
             //AlimChamps(True,RetCti)
          else
              begin
              AddLog(TraduireMemoire('--> Aucune fiche Tiers ou Contact correspondante'));
              AddLog(' ') ;
              RTCtiSelectionClient ( appelant, True);
              end;
          end;
      end;
{$IFNDEF CCATEST}
  end;
{$ENDIF}
end;

procedure TCCaForm.EffacerClick(Sender: TObject);
{$IFDEF CCATEST}
var pCall: IDataOfCall;
{$ENDIF}
begin
{$IFDEF CCATEST}
OnLineChangeCall(Sender,pCall);
{$ELSE}
Log.Lines.clear;
NOTELAPPELANT.Text:='';
TIERSAPPELANT.Text:='';
NATUREAPPELANT.Text:='';
LIBELLEAPPELANT.Text:= '';
NOTELAPPELE.Text:='';
TIERSAPPELE.Text:='';
LIBELLEAPPELE.Text:= '';
NATUREAPPELE.Text:='';
{$ENDIF}
end;

procedure TCCaForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AxApplication1.Visible:=AxALC_DEFAULT_VISIBLE;
end;

Function  TCCaForm.RTCtiSelectionClient ( appele : WideString; AffMul : Boolean ) : String;
var RetCti,StClient,StNat: string ;
begin
Repeat
CCAModeAppel := ModeAppelSortant;
V_PGI.ZoomOLE := true;  //Affichage en mode modal
Result := AGLLanceFiche ('RT', 'RTCTI_RECHTIERS','','',appele);
    V_PGI.ZoomOLE := false;
if Result<> '' then
    begin
    if not AppelEnCours then
        begin
        AlimChamps(False,Result);
        RetCti:=RTCtiDetailFicheClient(AUXIAPPELE.Text,NATUREAPPELE.Text) ;
        if RetCti <> '' then
            begin
            TIERSAPPELE.Text:=ReadTokenSt(RetCti);
            LIBELLEAPPELE.Text:= ReadTokenSt(RetCti);
            end;
        end
    else
        begin
        StNat:=ReadTokenSt(Result); StNat:=ReadTokenSt(Result); StNat:=ReadTokenSt(Result);
        StClient:=ReadTokenSt(Result);
        AGLLanceFiche('GC','GCTIERS','',StClient,'T_NATUREAUXI='+StNat) ;
        end;
    end;
until Result='';

end;

Function  TCCaForm.RTCtiDetailFicheClient ( StTiers , StNat : String ) : String;
var RetCti: string ;
begin
//AppelEnCours:=True ;
//Fo_CCA.AfficheFeu;
CCAModeAppel := ModeAppelSortant;
FicheMontee:=true;
RetCti:=RTCtiMonteFicheClient('NUMEROTER;CCA',StTiers,StNat) ;
FicheMontee:=false;
CtiRaccrocherFiche:=False;
end;

Function  TCCaForm.RTCtiReprendFicheClient ( TypeMontee, StTiers , StNat : String ) : String;
var RetCti: string ;
begin
FicheMontee:=true;
RetCti:=RTCtiMonteFicheClient(TypeMontee,StTiers,StNat) ;
FicheMontee:=false;
if (CtiRaccrocherFiche) or
   ( (CtiResultatAppel <> 0) and (CCAModeAppel = ModeAppelSortant)) then
    begin
    //AppelEnCours:=False ;
    ////AfficheFeu;
    //CtiRaccrocherFiche:=False ;
    //if (CtiResultatAppel <> 0) then CCAModeAppel := ModeAppelEntrant;
    end;
end;

procedure TCCaForm.AlimChamps( Appelant : Boolean; RetCti: String );
begin
if ( Appelant ) then
    begin
    TIERSAPPELANT.Text:=ReadTokenSt(RetCti);
    LIBELLEAPPELANT.Text:= ReadTokenSt(RetCti);
    NATUREAPPELANT.Text:=ReadTokenSt(RetCti);
    AUXIAPPELANT.Text:=ReadTokenSt(RetCti);
    NUMCONTACTAPPELANT.Text:=ReadTokenSt(RetCti);
    if NUMCONTACTAPPELANT.Text = '' then NUMCONTACTAPPELANT.Text:='0';
    NOMAPPELANT.Text:=ReadTokenSt(RetCti);
    end
else
    begin
    TIERSAPPELE.Text:=ReadTokenSt(RetCti);
    LIBELLEAPPELE.Text:= ReadTokenSt(RetCti);
    NATUREAPPELE.Text:=ReadTokenSt(RetCti);
    AUXIAPPELE.Text:=ReadTokenSt(RetCti);
    NUMCONTACTAPPELE.Text:=ReadTokenSt(RetCti);
    NOMAPPELE.Text:=ReadTokenSt(RetCti);
    end
end;

procedure TCCaForm.AfficheCoordonnees( Appelant : Boolean; Tiers,Libelle,Nature,Auxi,Num,Nom : String );
begin
if Fo_CCA.name<>'CCaForm' then
  begin
  if ( Appelant ) then
    begin
    TIERSAPPELANT.Text:=Tiers;
    LIBELLEAPPELANT.Text:= Libelle;
    NATUREAPPELANT.Text:=Nature;
    AUXIAPPELANT.Text:=Auxi;
    NUMCONTACTAPPELANT.Text:=Num;
    NOMAPPELANT.Text:=Nom;
    end
  else
    begin
    TIERSAPPELE.Text:=Tiers;
    LIBELLEAPPELE.Text:= Libelle;
    NATUREAPPELE.Text:=Nature;
    AUXIAPPELE.Text:=Auxi;
    NUMCONTACTAPPELE.Text:=Num;
    NOMAPPELE.Text:=Nom;
    end;
  end;
end;

procedure TCCaForm.AddLog (StLog : String);
begin
Log.Lines.add(StLog);
end;

procedure TCCaForm.ContactClick(Sender: TObject);
var Result,StCti,StNum,StNat,StClient : string;
begin
Repeat
Result:=AGLLanceFiche('RT','RTCTI_CONTACTS','','','');
if Result<> '' then    // Result= No contact + auxiliaire = nature auxi
    begin
    StNum:=ReadTokenSt(Result); StClient:=ReadTokenSt(Result); StNat:=ReadTokenSt(Result);
    if not AppelEnCours then
        begin
        //AlimChamps(False,Result);
        //AppelEnCours:=True ;
        //Fo_CCA.AfficheFeu;
        CCAModeAppel := ModeAppelSortant;
        FicheMontee:=true;
        StCti:=RTCtiMonteFicheContact('CTI;NUMEROTER;CCA',StClient,StrToInt(StNum));
        if StCti <> '' then
            Result:=ReadToKenPipe(StCti,'|')+';'+StNat+';'+StClient+';'+StNum+';'+ReadToKenPipe(StCti,'|');
        FicheMontee:=false;
        CtiRaccrocherFiche:=False;
        if (CtiResultatAppel <> 0) then
            begin
            //AppelEnCours:=False ;
            ////AfficheFeu;
            end;
        end
    else
        Result:=AGLLanceFiche('YY','YYCONTACT','T;'+StClient,StNum,'CTI') ;
    end;
Until Result='';
end;
procedure TCCaForm.TiersClick(Sender: TObject);
begin
AGLLanceFiche ('RT', 'RTCTI_RECHTIERS','' , '', '');
end;
procedure TCCaForm.ActionsClick(Sender: TObject);
begin
AGLLanceFiche ('RT', 'RTCTI_ACTIONS','' , '', '');
end;

end.



