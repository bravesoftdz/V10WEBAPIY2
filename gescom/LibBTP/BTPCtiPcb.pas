unit BTPCtiPcb;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {pcb_tlb,} StdCtrls, OleCtrls, Buttons,UtilCti,hctrls,comctrls, Mask,AGLInitGC,
  ExtCtrls, HTB97,ParamSoc,Utom,Entgc
{$IFDEF EAGLCLIENT}
      ,eFiche, Maineagl
{$ELSE}
      ,Fiche, FE_main
{$ENDIF}
,UTofTiersCti_Mul, PCb_TLB,RTRappel_Actions,HPanel,UIUtil, HEnt1, ImgList ;

Const
  Avance : integer = 1;
  Simple : integer = 0;
  ModeAppelSortant : integer = 1 ;
  ModeAppelEntrant : integer = 2 ;
  JournalCti : String = 'c:\pgi00\app\JournalCti.log';
type
  TResultatAppel =(raEnCours,raReponse,raOccupe,raRefus,raAbort,raFinCommunication);

  Tform1 = class(TForm)
    LOG: TMemo;
    APPELENTRANT: TGroupBox;
    NOTELAPPELANT: TEdit;
    Dock972: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    NATUREAPPELANT: TEdit;
    BSTOP: TToolbarButton97;
    RichEdit1: TRichEdit;
    NUMCONTACTAPPELANT: TEdit;
    ACTIONENTRANT: TGroupBox;
    NOMAPPELANT: TEdit;
    LNOMAPPELANT: TLabel;
    TIERSAPPELANT: TEdit;
    LTIERSAPPELANT: TLabel;
    LIBELLEAPPELANT: TEdit;
    ADRESSE1: TEdit;
    ADRESSE2: TEdit;
    ADRESSE3: TEdit;
    CODEPOSTAL: TEdit;
    VILLE: TEdit;
    LABDISPO: TLabel;
    DECROCHER: TToolbarButton97;
    ImgLTel: TImageList;
    FEU: TToolbarButton97;
    AUXIAPPELANT: TEdit;
    ControlPCB1: TControlPCB;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FermerClick(Sender: TObject);
    procedure AppelClick(Sender: TObject);
    procedure AttenteClick(Sender: TObject);
    procedure RepriseClick(Sender: TObject);
    procedure SerieAppelsClick(Sender: TObject);
    procedure ImprimerClick(Sender: TObject);
    procedure StopClick(Sender: TObject);
    procedure BoutonAppelClick(Sender: TObject);
    procedure DecrocherClick(Sender: TObject);
    procedure BoutonDecrocherClick(Sender: TObject);
    procedure RaccrocherClick(Sender: TObject);
    procedure BoutonAttenteClick(Sender: TObject);

    procedure ControlPCB1ResultatAppel(Sender: TObject; Resultat: Smallint;
      const NumeroTelephone: WideString);
    procedure ControlPCB1RequestIdentification(Sender: TObject;
      RefAppel: Integer; const appelant, appele: WideString);
    procedure ControlPCB1FinCommunicationTelephonique(Sender: TObject;
      const appelant, appele: WideString);
    procedure ControlPCB1AppelIdClientEntrant(Sender: TObject;
      const idclient, appele: WideString);
    procedure ControlPCB1AppelTelephonique(Sender: TObject; const appelant,
      appele: WideString);
    procedure ControlPCB1CommunicationTelephonique(Sender: TObject;
      const appelant, appele: WideString);
    procedure ControlPCB1EtatConnectionPCB(Sender: TObject;
      EtatPCB: WordBool);
    procedure ControlPCB1JournalAppelant(Sender: TObject; const appelant,
      appele: WideString);
    procedure ControlPCB1CommunicationIdClientEntrant(Sender: TObject;
      const idclient, appele: WideString);
    procedure ControlPCB1AppelEntrantInterrompu(Sender: TObject;
      RefAppel: Integer; const appelant: WideString);
    procedure ControlPCB1AppelEntrantRaccroche(Sender: TObject;
      RefAppel: Integer);
    procedure ControlPCB1AppelAbouti(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
    procedure ControlPCB1ErreurConnexion(Sender: TObject;
     RefAppel: Integer);
    procedure ControlPCB1CorrespondantEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
    procedure ControlPCB1RepriseCorrespondant(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
    procedure ControlPCB1AppelEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AfficheFeu;
    Function  RTCtiSelectionClient ( appele : WideString; AffMul : Boolean) : String;
    Function  RTCtiDetailFicheClient ( StTiers , StNat : String ) : String;
    Function  RTCtiReprendFicheClient ( TypeMontee, StTiers , StNat : String ) : String;
    Procedure RTCTINumeroteClient( Sender: TObject ; TelephoneAppele,StTiers,StNom,StNat,StAuxi : String );
    procedure AddLog (StLog : String);
    procedure AlimChamps( Appelant : Boolean; RetCti : String );
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);

    { Déclarations privées }
  private
    procedure RazChamps;

    { Déclarations publiques }
  public
    NumAppel : integer;
    NoTel : String;
		// true si appel en cours ou fiche client affichée, false sinon
    AppelEnCours : Boolean ;
    // pour gérer le raccrochage client
    FicheMontee : Boolean ;
  end;

Procedure RTLancePBC ;
procedure MonteMemoireCti;
procedure LiberememoireCti ;

var
    fo_PBC : Tform1;
    CtiErreurConnexion: Boolean;
    // etat appel sortant
    CtiResultatAppel : Smallint;
		// true si l'on a raccroché à partir de la fiche client
    // ou le correspondant a raccroché
    CtiRaccrocherFiche : Boolean ;
    CtiHeureDeb : TDateTime;
    // 1 = Appel sortant; 2 = Appel Entrant
    CtiModeAppel : Integer ;
    // No action crée
    CtiNumAct : Integer ;
    RTFormCti : TForm;
    FLogCti : textfile;
    CTIMounted : boolean;
implementation

{$R *.DFM}
Procedure RTLancePBC ;
Begin
	if not CTIMounted then exit;
	fo_PBC.Show;
end;

procedure MonteMemoireCti;
begin
	CTIMounted := false;
  if ( VH_GC.BTSeriaIntervention) and
  	 (GetParamSoc ('SO_BTCTIGESTION')) then
  begin
    AssignFile (FLogCti, JournalCti);
    Rewrite (FLogCti);
    TRY
    fo_PBC := Tform1.Create(Application);
    fo_PBC.DECROCHER.Down := False;
    CTIMounted := true;
    EXCEPT
     CTIMounted := false;
    END;
  end;

end;

procedure LiberememoireCti ;
begin
  if not CTIMounted then exit;
  CloseFile (FLogCti);
  fo_PBC.free;

end;

procedure Tform1.AfficheFeu;
begin

if AppelEnCours = True then
    begin
    FEU.ImageIndex := 2;
    LABDISPO.Caption:='Indisponible';
    end
else
    begin
    FEU.ImageIndex := 3;
    LABDISPO.Caption:='Disponible';
    // gérer manuellement : à tester RACCROCHER.Enabled:=False ;
    CtiErreurConnexion:=False;
    CtiModeAppel:=ModeAppelEntrant;
    CtiRaccrocherFiche:=False ;
    FicheMontee:=false;
    end
end;

Procedure Tform1.RTCTINumeroteClient( Sender: TObject ; TelephoneAppele,StTiers,StNom,StNat,StAuxi : String );
begin
AppelEnCours:=True ;
AfficheFeu;
CtiModeAppel := ModeAppelSortant;
Notel:=TelephoneAppele;
if Notel <> '' then
   begin
   //TIERSAPPELE.Text:= StTiers;
   //LIBELLEAPPELE.Text:= StNom;
   //NATUREAPPELE.Text:= StNat;
   //AUXIAPPELE.Text:=StAuxi;
   end;
AppelClick (Sender);
end;

procedure PGIBox( st, t : string);
begin
Application.MessageBox( pchar(st),pchar(t),MB_OK);
end;

procedure Tform1.FormShow(Sender: TObject);
begin
(*
ControlPCB1.startMode(Avance);
NoTel:='';
AppelEnCours:=False ;
//CtiRaccrocherFiche:=False ;
//FicheMontee:=false ;
RACCROCHER.Enabled:=False;
BATTENTE.Enabled:=False;
AfficheFeu;
NumAppel:=0;
CtiResultatAppel:=4;
CtiErreurConnexion:=False;
CtiHeureDeb:=0;
RTFormCti:=Nil;
*)
end;

procedure Tform1.FormDestroy(Sender: TObject);
begin
  if (ControlPCB1<> nil) then
  begin
    ControlPCB1.quit;
    ControlPcb1 := nil;
  end;
end;

procedure Tform1.FermerClick(Sender: TObject);
begin
	Hide;
end;

procedure Tform1.BoutonAppelClick(Sender: TObject);
var RetCti : string;
begin
{* CtiHeureDeb:=Time;
NoTel:=NOTELAPPELE.text;
if NoTel <> '' then
   RTCTINumeroteClient(Sender,NoTel,'','','','');
if NoTel = '' then
    begin // erreur connexion
    AppelEnCours:=False ;
    AfficheFeu;
    end
else
    begin // connexion établie
    FicheMontee:=true;
    CtiRaccrocherFiche:=False ;
    RetCti:=RTCtiAfficheClient ( NOTELAPPELE.text,'decroche;COUPLAGE');
    FicheMontee:=false;
    if (CtiRaccrocherFiche) and (CtiResultatAppel <> 0) then
        begin
        AppelEnCours:=False ;
        AfficheFeu;
        //CtiRaccrocherFiche:=False ;
        end;
    if RetCti <> ';;;' then
       AlimChamps(False,RetCti);
    end; *}
end;

procedure Tform1.AppelClick(Sender: TObject);
begin
{*
NOTELAPPELE.text:=NoTel;
if self.active then self.Show;
if NOTELAPPELE.text<>'' then
    begin
    DECROCHER.Enabled:=False ;
    APPEL.Enabled:=False ;
    RACCROCHER.Enabled:=True ;
    BATTENTE.Enabled:=True ;
    //composition du zéro devant le numéro
    if GetParamSoc('SO_BTCTILIGNEEXT') <> '' then
       NOTELAPPELE.text:=GetParamSoc('SO_BTCTILIGNEEXT')+ NOTELAPPELE.text;
    AddLog('--> Appel no '+ NOTELAPPELE.text);
    AddLog(' ') ;
    ControlPCB1.AppelUnClient(NOTELAPPELE.text) ;
    // si erreur connexion
    if CtiErreurConnexion then
        begin
        RACCROCHER.Enabled:=False ; BATTENTE.Enabled:=False ;
        NoTel:='';
        end;
    end;
    *}
end;

procedure Tform1.ImprimerClick(Sender: TObject);
begin
CloseFile (FLogCti);
RichEdit1.Lines.LoadFromFile(JournalCti);
RichEdit1.Print('Journal des appels en CTI');
Append (FLogCti);
end;

procedure Tform1.SerieAppelsClick(Sender: TObject);
begin
FicheMontee:=True;
RTLanceFiche_TiersCti_Mul ('RT', 'RTCTI_SERIETIERS','' , '', '');
AppelEnCours:=False ;
AfficheFeu;
end;

procedure Tform1.StopClick(Sender: TObject);
begin

if (BSTOP.down = True) then
   begin
   if (FEU.ImageIndex = 3) then
       begin
       FEU.ImageIndex := 2;
	     LABDISPO.Caption:='Indisponible';
       BImprimer.enabled:=false;
       Log.Lines.add(' ') ;
       AddLog('--> Connexion Téléphonie Interrompue ');
       AddLog(' ') ;
       if (ControlPCB1<> nil) then ControlPCB1.quit;
       end
   else
       BSTOP.down := False ;
   end
else
   begin
   FEU.ImageIndex := 3;
   LABDISPO.Caption:='Disponible';
   BImprimer.enabled:=true;
   ControlPCB1.startMode(Avance);
   end;
end;

procedure Tform1.DecrocherClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    AppelEnCours:=True ;
    AfficheFeu;
    CtiModeAppel:=ModeAppelEntrant;
    ControlPCB1.FctCstaAnswerCall(NumAppel);
    //RACCROCHER.Enabled:=True ;
    //BATTENTE.Enabled:=True ;
    //APPEL.Enabled:=False ;
    //DECROCHER.Enabled:=False ;
    end;
end;

procedure Tform1.BoutonDecrocherClick(Sender: TObject);
var RetCti 	 : string ;
var HeureFin : TDateTime;
begin

  if not DECROCHER.Down then
     Begin
     DECROCHER.ImageIndex := 1;
     DECROCHER.caption := 'Raccrocher';
     DECROCHER.Down := True;
     DecrocherClick(Sender);
     if NumAppel <> 0 then
        //RetCti:=RTCtiReprendFicheClient('COUPLAGE;decroche',AUXIAPPELANT.Text,NATUREAPPELANT.Text);
        RetCti:=RTCtiReprendFicheClient('COUPLAGE;decroche',NOTELAPPELANT.text,NATUREAPPELANT.Text);
		 end
  Else
     Begin
     DECROCHER.ImageIndex := 0;
     DECROCHER.caption := 'Décrocher';
     DECROCHER.Down := False;
		 AppelEnCours:=False ;
		 HeureFin:= Time;
		 HeureFin:=HeureFin-CtiHeureDeb ;
		 RaccrocherClick(Sender);
		 CtiHeureDeb:=0;
     RazChamps;
     end;

end;


procedure Tform1.RaccrocherClick(Sender: TObject);
begin

	if NumAppel <> 0 then
     begin
     ControlPCB1.FctCstaClearCall(NumAppel);
     NumAppel:=0;
     if AppelEnCours then
        CtiRaccrocherFiche := True
     else
        CtiRaccrocherFiche := False ;
     AfficheFeu;
     end;

end;

procedure Tform1.BoutonAttenteClick(Sender: TObject);
begin
{*
if BATTENTE.down = true then
   AttenteClick (Sender)
else
   RepriseClick (Sender);
*}
end;


procedure Tform1.AttenteClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    ControlPCB1.FctCstaHoldCall (NumAppel);
    //BATTENTE.down:= true;
    end;
end;

procedure Tform1.RepriseClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    ControlPCB1.FctCstaRetrieveCall  (NumAppel);
    //BATTENTE.down:= false;    
    end;
end;

procedure Tform1.ControlPCB1ResultatAppel(Sender: TObject;
  Resultat: Smallint; const NumeroTelephone: WideString);
begin
case TResultatAppel(Resultat) of
  raEnCours : AddLog('Appel en cours '+NumeroTelephone);
  raReponse : AddLog('Le correspondant a décroché');
  raOccupe  : AddLog('Le correspondant est occupé');
  raRefus   : AddLog('Appel refusé ( Cause de saturation , appel déjà en cours');
  raAbort   : AddLog('Appel interrompu par l’utilisateur');
  raFinCommunication : AddLog('Communication terminée');
end;
CtiResultatAppel:= Resultat ;
end;


procedure Tform1.ControlPCB1RequestIdentification(Sender: TObject;
  RefAppel: Integer; const appelant, appele: WideString);
  var nom,prenom,societe,RetCti : string;
begin

  NumAppel:=RefAppel;
  if CtiModeAppel = ModeAppelEntrant then
     NOTELAPPELANT.text:=appelant;

  AddLog('--> Demande d''identification Appel '+IntToStr(RefAppel));
  AddLog('    Appelant : '+appelant);
  AddLog('    Appelé   : '+appele);

  ControlPCB1.SendIdentificationPCB(RefAppel, Nom, Prenom, Societe);

  if CtiModeAppel = ModeAppelEntrant then
      begin
      //DECROCHER.Enabled:=True ;
      AppelEnCours:=True ;
      AfficheFeu;
      FicheMontee:=true;
      CtiRaccrocherFiche:=False ;
      RazChamps;
      CtiHeureDeb:=Time;
      RetCti:=RTCtiAfficheClient( appelant,'appel;COUPLAGE');
      FicheMontee:=False;
      if (CtiRaccrocherFiche) then
          begin
          AppelEnCours:=False ;
          AfficheFeu;
          end;

      if RetCti <> '' then
         AlimChamps(True,RetCti)
      else
         begin
         AddLog('--> Aucune fiche Tiers ou Contact correspondante');
         AddLog(' ') ;
         PGIBox('Aucune fiche Tiers ou Contact correspondante', 'Erreur Téléphone');
         end;
      end;
end;

procedure Tform1.ControlPCB1FinCommunicationTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog('--> Fin communication ');
AddLog('    Appelant : '+appelant);
AddLog('    Appelé   : '+appele);

end;

procedure Tform1.ControlPCB1AppelIdClientEntrant(Sender: TObject;
  const idclient, appele: WideString);
begin
AddLog('--> Appel ID Entrant ');
AddLog('    IdClient : '+idclient);
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1AppelTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog('--> Appel telephonique ');
AddLog('    Appelant : '+appelant);
AddLog('    Appelé   : '+appele);
if not (self.Active) then self.show; 
end;

procedure Tform1.ControlPCB1CommunicationTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog('--> Début Communication Telephonique ');
AddLog('    Appelant : '+appelant);
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1EtatConnectionPCB(Sender: TObject;
  EtatPCB: WordBool);
begin
AddLog('--> Connexion Téléphonie établie ');
AddLog(' ') ;
end;

procedure Tform1.ControlPCB1JournalAppelant(Sender: TObject;
  const appelant, appele: WideString);
var HeureFin : TDateTime;
St:string;
begin
AddLog('--> Votre correspondant a raccroché ');
AddLog('    Appelant : '+appelant);
AddLog('    Appelé   : '+appele);
if RTFormCti <> Nil then St:=RTFormCti.Name;
// génération d'une action si le tiers appelant existe et raccroche sans que l'on ait déccrocher
HeureFin:=Time;
// par sécurité mais normalement ds ce cas, toujours appel entrant
if (CtiModeAppel = ModeAppelEntrant ) and (TIERSAPPELANT.Text <> '') then
   RTCTIGenereAction(False,TIERSAPPELANT.Text,AUXIAPPELANT.Text,CtiHeureDeb,HeureFin,'',CtiModeAppel,StrToInt(NUMCONTACTAPPELANT.Text));
end;

procedure Tform1.ControlPCB1AppelEntrantRaccroche(Sender: TObject;
  RefAppel: Integer);
//var OM : TOM ;
// il s'avere que cet évênement arrive également sur un raccrochage de l'utilisateur !
begin
AddLog('--> Appel Entrant Raccroché');
AddLog('    Id Appel   : '+IntToStr(RefAppel));


////////// essai
if RTFormCti <> Nil then
    begin
    if (RTFormCti is TFFiche) then
        begin
        //OM:=TFFiche(RTFormCti).OM ;
        //if (OM is TOM_TIERS) then
            //TOM_TIERS(OM).RTCTIAfficheMessageCTI;
        end;
    end;

////////// fin essai
CtiRaccrocherFiche := True  ;
if FicheMontee=false then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;
end;

procedure Tform1.ControlPCB1CommunicationIdClientEntrant(Sender: TObject;
  const idclient, appele: WideString);
begin
AddLog('--> Communication ID Entrant ');
AddLog('    IdClient : '+idclient);
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1AppelEntrantInterrompu(Sender: TObject;
  RefAppel: Integer; const appelant: WideString);
begin
AddLog('--> Appel Entrant Interrompu');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
AddLog('    Appelant   : '+appelant);
CtiRaccrocherFiche := True  ;
if FicheMontee=false then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;
end;

procedure Tform1.ControlPCB1AppelAbouti(Sender: TObject; RefAppel: Integer;
 const appele: WideString);
begin
AddLog('--> Appel Sortant Abouti');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1ErreurConnexion(Sender: TObject;
  RefAppel: Integer);
begin
AddLog('--> Erreur sur la connexion');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
CtiErreurConnexion:=True;
end;

procedure Tform1.ControlPCB1CorrespondantEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog('--> Correspondant en Attente');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1RepriseCorrespondant(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog('--> Reprise du Correspondant');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
AddLog('    Appelé   : '+appele);
end;

procedure Tform1.ControlPCB1AppelEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog('--> Appel en Attente');
AddLog('    Id Appel   : '+IntToStr(RefAppel));
AddLog('    Appelant   : '+appele);
end;
                  

procedure Tform1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//CloseFile (FLogCti);
	if (ControlPCB1<> nil) then
  begin
  	ControlPCB1.quit;
    ControlPcb1 := nil;
  end;
end;

Function  Tform1.RTCtiSelectionClient ( appele : WideString; AffMul : Boolean ) : String;
var RetCti,StClient,StNat: string ;
begin
Repeat
Result := AGLLanceFiche ('RT', 'RTCTI_RECHTIERS','','','');

if Result<> '' then
    begin
    if not AppelEnCours then
        begin
        AlimChamps(False,Result);
        //RetCti:=RTCtiDetailFicheClient(AUXIAPPELE.Text,NATUREAPPELE.Text) ;
        if RetCti <> '' then
            begin
            //TIERSAPPELE.Text:=ReadTokenSt(RetCti);
            //LIBELLEAPPELE.Text:= ReadTokenSt(RetCti);
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

Function  Tform1.RTCtiDetailFicheClient ( StTiers , StNat : String ) : String;
var RetCti: string ;
begin
AppelEnCours:=True ;
fo_PBC.AfficheFeu;
CtiModeAppel := ModeAppelSortant;
FicheMontee:=true;
RetCti:=RTCtiMonteFicheClient('NUMEROTER',StTiers,StNat) ;
FicheMontee:=false;
CtiRaccrocherFiche:=False;
if (CtiResultatAppel <> 0) then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;
end;

Function  Tform1.RTCtiReprendFicheClient ( TypeMontee, StTiers , StNat : String ) : String;
var RetCti: string ;
begin
FicheMontee:=true;
RetCti:=RTCtiMonteFicheClient(TypeMontee,StTiers,StNat) ;
FicheMontee:=false;
if (CtiRaccrocherFiche) or
   ( (CtiResultatAppel <> 0) and (CtiModeAppel = ModeAppelSortant)) then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;
end;

Procedure TForm1.RazChamps;
Begin
  NOTELAPPELANT.Text := '';
  NOMAPPELANT.Text := '';
  TIERSAPPELANT.Text := '';
  LIBELLEAPPELANT.Text := '';
  ADRESSE1.Text := '';
	ADRESSE2.Text := '';
	ADRESSE3.Text := '';
	CODEPOSTAL.Text := '';
	VILLE.Text := '';
End;

procedure Tform1.AlimChamps( Appelant : Boolean; RetCti: String );
begin
  NOTELAPPELANT.Text:=ReadTokenSt(RetCti);
  NOMAPPELANT.text:=ReadTokenSt(RetCti);
  AUXIAPPELANT.text:=ReadTokenSt(RetCti);
  TIERSAPPELANT.Text:=ReadTokenSt(RetCti);
  LIBELLEAPPELANT.Text:= ReadTokenSt(RetCti);
  NATUREAPPELANT.Text:=ReadTokenSt(RetCti);
  ADRESSE1.Text:=ReadTokenSt(RetCti);
  ADRESSE2.Text:=ReadTokenSt(RetCti);
  ADRESSE3.Text:=ReadTokenSt(RetCti);
  CODEPOSTAL.Text:=ReadTokenSt(RetCti);
  VILLE.Text:=ReadTokenSt(RetCti);
end;

procedure Tform1.AddLog (StLog : String);
begin
  Log.Lines.add(StLog);
  writeln (FLogCti,StLog);
end;

procedure Tform1.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure Tform1.FormCreate(Sender: TObject);
begin
  ControlPCB1.startMode(Avance);

  NoTel:='';
  AppelEnCours:=False ;
  //CtiRaccrocherFiche:=False ;
  //FicheMontee:=false ;
  AfficheFeu;
  NumAppel:=0;
  CtiResultatAppel:=4;
  CtiErreurConnexion:=False;
  CtiHeureDeb:=0;
  RTFormCti:=Nil;
end;

end.

