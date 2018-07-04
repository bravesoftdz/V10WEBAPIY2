unit CtiPcb;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms,
  StdCtrls, Buttons,UtilCti,hctrls,comctrls,
  ExtCtrls, HTB97,ParamSoc
{$IFDEF EAGLCLIENT}
      ,eFiche, Maineagl
{$ELSE}
      ,Fiche, FE_main
{$ENDIF}
,UTofTiersCti_Mul, PCb_TLB,RTRappel_Actions,HEnt1, Graphics, Mask, OleCtrls, CBPPath
 ;

Const
  Avance : integer = 1;
  Simple : integer = 0;
  ModeAppelSortant : integer = 1 ;
  ModeAppelEntrant : integer = 2 ;

type
  TResultatAppel =(raEnCours,raReponse,raOccupe,raRefus,raAbort,raFinCommunication);

  TForm1 = class(TForm)
    LOG: TMemo;
    APPELENTRANT: TGroupBox;
    NOTELAPPELANT: TEdit;
    LNOTELAPPELANT: TLabel;
    LTIERSAPPELANT: TLabel;
    TIERSAPPELANT: TEdit;
    LIBELLEAPPELANT: TEdit;
    ACTIONENTRANT: TGroupBox;
    DECROCHER: THSpeedButton;
    RACCROCHER: THSpeedButton;
    APPELSORTANT: TGroupBox;
    LNOTELAPPELE: TLabel;
    NOTELAPPELE: TEdit;
    LTIERSAPPELE: TLabel;
    LIBELLEAPPELE: TEdit;
    LLIBELLEAPPELE: TLabel;
    TIERSAPPELE: THCritMaskEdit;
    Dock972: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    NATUREAPPELE: TEdit;
    NATUREAPPELANT: TEdit;
    ZOOMAPPELANT: THBitBtn;
    APPEL: THBitBtn;
    ZOOMAPPELE: THBitBtn;
    FEUROUGE: TImage;
    FEUVERT: TImage;
    INDISPONIBLE: TLabel;
    DISPONIBLE: TLabel;
    BSTOP: TToolbarButton97;
    AUXIAPPELANT: TEdit;
    AUXIAPPELE: TEdit;
    RichEdit1: TRichEdit;
    SERIEAPPELS: THBitBtn;
    BATTENTE: THSpeedButton;
    BCONTACT: TToolbarButton97;
    BTIERS: TToolbarButton97;
    EFFACER: TToolbarButton97;
    BACTIONS: TToolbarButton97;
    NOMAPPELANT: TEdit;
    LNOMAPPELANT: TLabel;
    NOMAPPELE: TEdit;
    LNOMAPPELE: TLabel;
    NUMCONTACTAPPELE: TEdit;
    NUMCONTACTAPPELANT: TEdit;
    ControlPCB1: TControlPCB;
    LLIBELLEAPPELANT: TLabel;
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
    procedure AppelListeTiers(Sender: TObject);
    procedure DecrocherClick(Sender: TObject);
    procedure BoutonDecrocherClick(Sender: TObject);
    procedure RaccrocherClick(Sender: TObject);
    procedure BoutonRaccrocherClick(Sender: TObject);
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

    procedure EffacerClick(Sender: TObject);
    procedure ZoomClick(Sender: TObject);

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AfficheFeu;
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
    procedure BAideClick(Sender: TObject);
    { Déclarations privées }
  public
    { Déclarations publiques }
    NumAppel : integer;
    NoTel : String;
    AppelEnCours : Boolean ; // true si appel en cours ou fiche client affichée, false sinon
    FicheMontee : Boolean ; // pour gérer le raccrochage client
    FLogCti : textfile;
  end;
Procedure RTLancePBC ;
var
    Fo_PBC : TForm1;
    CtiErreurConnexion: Boolean;
    CtiResultatAppel : Smallint; // etat appel sortant
    CtiRaccrocherFiche : Boolean ; // true si l'on a raccroché à partir de la fiche client ou le correspondant a raccroché
    CtiHeureDeb : TDateTime;
    CtiModeAppel : Integer ; // 1 = Appel sortant; 2 = Appel Entrant
    CtiNumAct : Integer ; // No action crée
    RTFormCti : TForm;
    NoTel,JournalCti : String;
implementation

{$R *.DFM}
Procedure RTLancePBC ;
Begin
//     JournalCti := ExtractFilePath(Application.ExeName)+'JournalCti.log';
     JournalCti := IncludeTrailingBackslash (TCBPPath.GetCegidDataDistri) +'JournalCti.log';
     Fo_PBC := TForm1.Create (Application);
     RTEndThreadActions();
     if Assigned(Fo_PBC) then
        Fo_PBC.ShowModal;
end;

procedure TForm1.AfficheFeu;
begin

if AppelEnCours = True then
    begin
    FEUVERT.visible:=false;
    DISPONIBLE.visible:=false;
    FEUROUGE.visible:=true;
    INDISPONIBLE.visible:=true;
    end
else
    begin
    FEUVERT.visible:=true;
    DISPONIBLE.visible:=true;
    FEUROUGE.visible:=false;
    INDISPONIBLE.visible:=false;
    // gérer manuellement : à tester RACCROCHER.Enabled:=False ;
    DECROCHER.Enabled:=False ;
    APPEL.Enabled:=True ;
    CtiErreurConnexion:=False;
    CtiModeAppel:=ModeAppelEntrant;
    CtiRaccrocherFiche:=False ;
    FicheMontee:=false;
    end
end;

Procedure TForm1.RTCTINumeroteClient( Sender: TObject ; TelephoneAppele,StTiers,StNom,StNat,StAuxi : String );
begin
AppelEnCours:=True ;
AfficheFeu;
CtiModeAppel := ModeAppelSortant;
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

procedure PGIBox( st, t : string);
begin
Application.MessageBox( pchar(st),pchar(t),MB_OK);
end;

procedure TForm1.FormShow(Sender: TObject);
begin
AssignFile (FLogCti, JournalCti);
Rewrite (FLogCti);
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
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
if (ControlPCB1<> nil) then ControlPCB1.quit;
end;

procedure TForm1.FermerClick(Sender: TObject);
begin
Close;
end;

procedure TForm1.BoutonAppelClick(Sender: TObject);
var RetCti : string;
begin
CtiHeureDeb:=Time;
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
    //RTCTIGenereAction(True,TIERSAPPELE.Text,AUXIAPPELE.Text,CtiHeureDeb,HeureFin,'',CtiModeAppel)

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
    end;
end;

procedure TForm1.AppelClick(Sender: TObject);
begin
NOTELAPPELE.text:=NoTel;
if NOTELAPPELE.text<>'' then
    begin
    DECROCHER.Enabled:=False ;
    APPEL.Enabled:=False ;
    RACCROCHER.Enabled:=True ;
    BATTENTE.Enabled:=True ;
    { composition du zéro devant le numéro }
    if GetParamsocSecur('SO_RTCTILIGNEEXT','00') <> '' then
       NOTELAPPELE.text:=GetParamsocSecur('SO_RTCTILIGNEEXT','00')+NOTELAPPELE.text;
    AddLog(TraduireMemoire('--> Appel no ')+NOTELAPPELE.text);
    AddLog(' ') ;
    ControlPCB1.AppelUnClient(NOTELAPPELE.text) ;
    // si erreur connexion
    {if CtiErreurConnexion then
        begin
        NoTel:='';
        AppelEnCours:=False ;
        end;
    AfficheFeu;}
    if CtiErreurConnexion then
        begin
        RACCROCHER.Enabled:=False ; BATTENTE.Enabled:=False ;
        NoTel:='';
        end;
    end;
end;

procedure TForm1.ImprimerClick(Sender: TObject);
begin
CloseFile (FLogCti);
//RichEdit1.Create (Sender);
RichEdit1.Lines.LoadFromFile(JournalCti);
RichEdit1.Print(TraduireMemoire('Journal des appels en CTI'));
//RichEdit1.Destroy ;
Append (FLogCti);
end;

procedure TForm1.SerieAppelsClick(Sender: TObject);
begin
FicheMontee:=True;
RTLanceFiche_TiersCti_Mul ('RT', 'RTCTI_SERIETIERS','' , '', '');
//FicheMontee:=False;
AppelEnCours:=False ;
AfficheFeu;
end;

procedure TForm1.StopClick(Sender: TObject);
begin
if (BSTOP.down = True) then
   begin
   if (FEUVERT.visible = true) then
       begin
       FEUVERT.visible:=False;
       FEUROUGE.visible:=True;
       INDISPONIBLE.visible:=true;
       DISPONIBLE.visible:=false;
       ZOOMAPPELANT.enabled:=false;
       ZOOMAPPELE.enabled:=false;
       APPEL.enabled:=false;
       BImprimer.enabled:=false;
       EFFACER.enabled:=false;
       SERIEAPPELS.enabled:=false;
       Log.Lines.add(' ') ;
       AddLog(TraduireMemoire('--> Connexion Téléphonie Interrompue '));
       AddLog(' ') ;

       if (ControlPCB1<> nil) then ControlPCB1.quit;
       end
   else
       BSTOP.down := False ;
   end
else
   begin
   FEUVERT.visible:=True;
   FEUROUGE.visible:=False;
   INDISPONIBLE.visible:=false;
   DISPONIBLE.visible:=true;
   ZOOMAPPELANT.enabled:=true;
   ZOOMAPPELE.enabled:=true;
   APPEL.enabled:=true;
   BImprimer.enabled:=true;
   EFFACER.enabled:=true;
   SERIEAPPELS.enabled:=true;
   ControlPCB1.startMode(Avance);
   end;
end;

procedure TForm1.AppelListeTiers(Sender: TObject);
begin
//NOTELAPPELE.text:=DispatchRecherche(TIERSAPPELE,2,'','','');

RTCtiSelectionClient ( TIERSAPPELE.text, True);

if CtiRaccrocherFiche then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    //CtiRaccrocherFiche:=False ;
    end;

end;

procedure TForm1.DecrocherClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    AppelEnCours:=True ;
    AfficheFeu;
    CtiModeAppel:=ModeAppelEntrant;
    ControlPCB1.FctCstaAnswerCall(NumAppel);
    RACCROCHER.Enabled:=True ;
    BATTENTE.Enabled:=True ;    
    APPEL.Enabled:=False ;
    DECROCHER.Enabled:=False ;
    end;
end;

procedure TForm1.BoutonDecrocherClick(Sender: TObject);
var RetCti : string ;
begin
DecrocherClick(Sender);
if NumAppel <> 0 then
    RetCti:=RTCtiReprendFicheClient('COUPLAGE;decroche',AUXIAPPELANT.Text,NATUREAPPELANT.Text);
end;

procedure TForm1.BoutonRaccrocherClick(Sender: TObject);
var HeureFin : TDateTime;
begin
AppelEnCours:=False ;
HeureFin:= Time;
HeureFin:=HeureFin-CtiHeureDeb ;
RaccrocherClick(Sender);
if GetParamsocSecur('SO_RTCTIGENACTEFFECT',false) then
    if CtiModeAppel=ModeAppelSortant then
       RTCTIMajDureeAction (AUXIAPPELE.Text,CtiNumAct,HeureFin)
       //RTCTIGenereAction(True,TIERSAPPELE.Text,AUXIAPPELE.Text,CtiHeureDeb,HeureFin,'',CtiModeAppel)
    else
       if TIERSAPPELANT.Text <> '' then
          RTCTIMajDureeAction (AUXIAPPELANT.Text,CtiNumAct,HeureFin);
          //RTCTIGenereAction(True,TIERSAPPELANT.Text,AUXIAPPELANT.Text,CtiHeureDeb,HeureFin,'',CtiModeAppel);

CtiHeureDeb:=0;
end;

procedure TForm1.RaccrocherClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    ControlPCB1.FctCstaClearCall(NumAppel);
    NumAppel:=0;
    if AppelEnCours then
        CtiRaccrocherFiche := True
    else
        CtiRaccrocherFiche := False ;
    { déjà dans affichefeu
    DECROCHER.Enabled:=False ;
    APPEL.Enabled:=True ;
    RACCROCHER.Enabled:=False ;}
    AfficheFeu;
    end;
RACCROCHER.Enabled:=False ;
BATTENTE.Enabled:=False ;
end;

procedure TForm1.BoutonAttenteClick(Sender: TObject);
begin
if BATTENTE.down = true then
   AttenteClick (Sender)
else
   RepriseClick (Sender);
end;


procedure TForm1.AttenteClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    ControlPCB1.FctCstaHoldCall (NumAppel);
    BATTENTE.down:= true;
    end;
end;

procedure TForm1.RepriseClick(Sender: TObject);
begin
if NumAppel <> 0 then
    begin
    ControlPCB1.FctCstaRetrieveCall  (NumAppel);
    BATTENTE.down:= false;    
    end;
end;

procedure TForm1.ControlPCB1ResultatAppel(Sender: TObject;
  Resultat: Smallint; const NumeroTelephone: WideString);
begin
case TResultatAppel(Resultat) of
  raEnCours : AddLog(TraduireMemoire('Appel en cours ')+NumeroTelephone);
  raReponse : AddLog(TraduireMemoire('Le correspondant a décroché'));
  raOccupe  : AddLog(TraduireMemoire('Le correspondant est occupé'));
  raRefus   : AddLog(TraduireMemoire('Appel refusé ( Cause de saturation , appel déjà en cours'));
  raAbort   : AddLog(TraduireMemoire('Appel interrompu par l’utilisateur'));
  raFinCommunication : AddLog(TraduireMemoire('Communication terminée'));
end;
CtiResultatAppel:= Resultat ;
end;


procedure TForm1.ControlPCB1RequestIdentification(Sender: TObject;
  RefAppel: Integer; const appelant, appele: WideString);
  var nom,prenom,societe,RetCti : string;
begin
  NumAppel:=RefAppel;
  if CtiModeAppel = ModeAppelEntrant then
     NOTELAPPELANT.text:=appelant;
  AddLog(TraduireMemoire('--> Demande d''identification Appel ')+IntToStr(RefAppel));
  AddLog(TraduireMemoire('    Appelant : ')+appelant);
  AddLog(TraduireMemoire('    Appelé   : ')+appele);
  {if appelant='0146548712' then
     begin
     Nom:='MARTIN';Prenom:='Pierre';Societe:='TREDOC INDUSTRIE';
     end else
     begin
     Nom:='DECHAMPS';Prenom:='Jean-François';Societe:='Fromagerie BEL';
     end; }
  ControlPCB1.SendIdentificationPCB(RefAppel, Nom, Prenom, Societe);
  if CtiModeAppel = ModeAppelEntrant then
      begin
      DECROCHER.Enabled:=True ;
//      RACCROCHER.Enabled:=False ;
      AppelEnCours:=True ;
      AfficheFeu;
      FicheMontee:=true;
      APPEL.Enabled:=False ;
      CtiRaccrocherFiche:=False ;
      TIERSAPPELANT.Text:='';
      NATUREAPPELANT.Text:='';
      LIBELLEAPPELANT.Text:= '';
      CtiHeureDeb:=Time;
      RetCti:=RTCtiAfficheClient ( appelant,'appel;COUPLAGE');
      FicheMontee:=False;
      if (CtiRaccrocherFiche) {and (CtiResultatAppel <> 0)} then
          begin
          AppelEnCours:=False ;
          AfficheFeu;
          //CtiRaccrocherFiche:=False ;
          end;

      if RetCti <> ';;;' then
         AlimChamps(True,RetCti)
      else
          // pas de fiche tiers, on recherche une fiche contact
          begin
          RetCti:=RTCtiAfficheContact ( appelant,'appel;COUPLAGE;CTI');
          FicheMontee:=False;
          if (CtiRaccrocherFiche) {and (CtiResultatAppel <> 0)} then
              begin
              AppelEnCours:=False ;
              AfficheFeu;
              //CtiRaccrocherFiche:=False ;
              end;

          if RetCti <> ';;;;0' then
             //AlimChamps(True,RetCti)
          else
              begin
              AddLog(TraduireMemoire('--> Aucune fiche Tiers ou Contact correspondante'));
              AddLog(' ') ;
              end;
          end;
      end;
end;

procedure TForm1.ControlPCB1FinCommunicationTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog(TraduireMemoire('--> Fin communication '));
AddLog(TraduireMemoire('    Appelant : ')+appelant);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
//AfficheFeu;
end;

procedure TForm1.ControlPCB1AppelIdClientEntrant(Sender: TObject;
  const idclient, appele: WideString);
begin
AddLog(TraduireMemoire('--> Appel ID Entrant '));
AddLog(TraduireMemoire('    IdClient : ')+idclient);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1AppelTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog(TraduireMemoire('--> Appel téléphonique '));
AddLog(TraduireMemoire('    Appelant : ')+appelant);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1CommunicationTelephonique(Sender: TObject;
  const appelant, appele: WideString);
begin
AddLog(TraduireMemoire('--> Début Communication Téléphonique '));
AddLog(TraduireMemoire('    Appelant : ')+appelant);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1EtatConnectionPCB(Sender: TObject;
  EtatPCB: WordBool);
begin
AddLog(TraduireMemoire('--> Connexion Téléphonie établie '));
AddLog(' ') ;
end;

procedure TForm1.ControlPCB1JournalAppelant(Sender: TObject;
  const appelant, appele: WideString);
var HeureFin : TDateTime;
St:string;
begin
AddLog(TraduireMemoire('--> Votre correspondant a raccroché '));
AddLog(TraduireMemoire('    Appelant : ')+appelant);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
if RTFormCti <> Nil then St:=RTFormCti.Name;
{ à priori inutile puisque la fonction suivante est également activée
CtiRaccrocherFiche := True  ;
if FicheMontee=false then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;}
//CtiModeAppel:=ModeAppelEntrant;

// génération d'une action si le tiers appelant existe et raccroche sans que l'on ait déccrocher
HeureFin:=Time;
// par sécurité mais normalement ds ce cas, toujours appel entrant
if (CtiModeAppel = ModeAppelEntrant ) and (TIERSAPPELANT.Text <> '') then
   RTCTIGenereAction(False,TIERSAPPELANT.Text,AUXIAPPELANT.Text,CtiHeureDeb,HeureFin,'',CtiModeAppel,ValeurI(NUMCONTACTAPPELANT.Text));
end;

procedure TForm1.ControlPCB1AppelEntrantRaccroche(Sender: TObject;
  RefAppel: Integer);
//var OM : TOM ;
// il s'avere que cet évênement arrive également sur un raccrochage de l'utilisateur !
begin
AddLog(TraduireMemoire('--> Appel Entrant Raccroché'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));


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
//CtiModeAppel:=ModeAppelEntrant;
end;

procedure TForm1.ControlPCB1CommunicationIdClientEntrant(Sender: TObject;
  const idclient, appele: WideString);
begin
AddLog(TraduireMemoire('--> Communication ID Entrant '));
AddLog(TraduireMemoire('    IdClient : ')+idclient);
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1AppelEntrantInterrompu(Sender: TObject;
  RefAppel: Integer; const appelant: WideString);
begin
AddLog(TraduireMemoire('--> Appel Entrant Interrompu'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
AddLog(TraduireMemoire('    Appelant   : ')+appelant);
CtiRaccrocherFiche := True  ;
if FicheMontee=false then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    end;
//CtiModeAppel:=ModeAppelEntrant;
end;

procedure TForm1.ControlPCB1AppelAbouti(Sender: TObject; RefAppel: Integer;
 const appele: WideString);
begin
AddLog(TraduireMemoire('--> Appel Sortant Abouti'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1ErreurConnexion(Sender: TObject;
  RefAppel: Integer);
begin
AddLog(TraduireMemoire('--> Erreur sur la connexion'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
CtiErreurConnexion:=True;
end;

procedure TForm1.ControlPCB1CorrespondantEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog(TraduireMemoire('--> Correspondant en Attente'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1RepriseCorrespondant(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog(TraduireMemoire('--> Reprise du Correspondant'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
AddLog(TraduireMemoire('    Appelé   : ')+appele);
end;

procedure TForm1.ControlPCB1AppelEnAttente(Sender: TObject; RefAppel: Integer;
     const appele: WideString);
begin
AddLog(TraduireMemoire('--> Appel en Attente'));
AddLog(TraduireMemoire('    Id Appel   : ')+IntToStr(RefAppel));
AddLog(TraduireMemoire('    Appelant   : ')+appele);
end;

procedure TForm1.EffacerClick(Sender: TObject);
begin
Log.Lines.clear;
NOTELAPPELANT.Text:='';
TIERSAPPELANT.Text:='';
NATUREAPPELANT.Text:='';
LIBELLEAPPELANT.Text:= '';
NOTELAPPELE.Text:='';
TIERSAPPELE.Text:='';
LIBELLEAPPELE.Text:= '';
NATUREAPPELE.Text:='';
end;

procedure TForm1.ZoomClick(Sender: TObject);
var RetCti : string;
begin
CtiRaccrocherFiche:=False;
if TBitBtn(Sender).Name = 'ZOOMAPPELANT' then
    begin
    if (not AppelEnCours) then
        begin
        CtiModeAppel:=ModeAppelSortant;
        if (TIERSAPPELANT.Text <> '') then
            RTCtiDetailFicheClient (AUXIAPPELANT.Text,NATUREAPPELANT.Text);
        end
    else
        if DECROCHER.Enabled=False then // Appel en cours
            if BATTENTE.down = true then
               RTCtiReprendFicheClient('COUPLAGE;attente',AUXIAPPELANT.Text,NATUREAPPELANT.Text)
            else
               RTCtiReprendFicheClient('COUPLAGE;decroche',AUXIAPPELANT.Text,NATUREAPPELANT.Text)
        else
            // Appel entrant en attente
            RTCtiReprendFicheClient('COUPLAGE;appel',AUXIAPPELANT.Text,NATUREAPPELANT.Text);
    end
else
    begin
    if (not AppelEnCours) then
        begin
        CtiModeAppel:=ModeAppelSortant;
        if (TIERSAPPELE.Text <> '') then
            RTCtiDetailFicheClient (AUXIAPPELE.Text,NATUREAPPELE.Text)
        else
        if (NOTELAPPELE.Text <> '') then // Detail fiche qu'à partir du numéro pourquoi pas ?
            begin
            AppelEnCours:=True ;
            AfficheFeu;
            FicheMontee:=true;
            APPEL.Enabled:=False ;
            CtiRaccrocherFiche:=False ;
            RetCti:=RTCtiAfficheClient ( NOTELAPPELE.Text,'NUMEROTER');
            FicheMontee:=False;
            if (CtiRaccrocherFiche) or (CtiResultatAppel <> 0) then
                begin
                AppelEnCours:=False ;
                AfficheFeu;
                //CtiRaccrocherFiche:=False ;
                //if (CtiResultatAppel <> 0) then CtiModeAppel := ModeAppelEntrant;
                end;

            if RetCti <> ';;;' then
               AlimChamps(False,RetCti);
            end;
        end
    else
        if DECROCHER.Enabled=False then // Appel en cours
            if BATTENTE.down = true then
               RTCtiReprendFicheClient('COUPLAGE;attente',AUXIAPPELE.Text,NATUREAPPELE.Text)
            else
               RTCtiReprendFicheClient('COUPLAGE;decroche',AUXIAPPELE.Text,NATUREAPPELE.Text)
        else
            // Appel entrant en attente
            RTCtiReprendFicheClient('COUPLAGE;appel',AUXIAPPELE.Text,NATUREAPPELE.Text);
    end
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
CloseFile (FLogCti);
if (ControlPCB1<> nil) then ControlPCB1.quit;
end;

Function  TForm1.RTCtiSelectionClient ( appele : WideString; AffMul : Boolean ) : String;
var RetCti,StClient,StNat: string ;
begin
Repeat
Result := AGLLanceFiche ('RT', 'RTCTI_RECHTIERS','','','');

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

Function  TForm1.RTCtiDetailFicheClient ( StTiers , StNat : String ) : String;
var RetCti: string ;
begin
AppelEnCours:=True ;
Fo_PBC.AfficheFeu;
CtiModeAppel := ModeAppelSortant;
FicheMontee:=true;
RetCti:=RTCtiMonteFicheClient('NUMEROTER',StTiers,StNat) ;
FicheMontee:=false;
CtiRaccrocherFiche:=False;
if {(CtiRaccrocherFiche) or }(CtiResultatAppel <> 0) then
    begin
    AppelEnCours:=False ;
    AfficheFeu;
    //CtiRaccrocherFiche:=False ;
    //if (CtiResultatAppel <> 0) then CtiModeAppel := ModeAppelEntrant;
    end;
end;

Function  TForm1.RTCtiReprendFicheClient ( TypeMontee, StTiers , StNat : String ) : String;
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
    //CtiRaccrocherFiche:=False ;
    //if (CtiResultatAppel <> 0) then CtiModeAppel := ModeAppelEntrant;
    end;
end;

procedure TForm1.AlimChamps( Appelant : Boolean; RetCti: String );
begin
if ( Appelant ) then
    begin
    TIERSAPPELANT.Text:=ReadTokenSt(RetCti);
    LIBELLEAPPELANT.Text:= ReadTokenSt(RetCti);
    NATUREAPPELANT.Text:=ReadTokenSt(RetCti);
    AUXIAPPELANT.Text:=ReadTokenSt(RetCti);
    NUMCONTACTAPPELANT.Text:=ReadTokenSt(RetCti);
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

procedure TForm1.AfficheCoordonnees( Appelant : Boolean; Tiers,Libelle,Nature,Auxi,Num,Nom : String );
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
    end
end;

procedure TForm1.AddLog (StLog : String);
begin
Log.Lines.add(StLog);
writeln (FLogCti,StLog);
end;

procedure TForm1.ContactClick(Sender: TObject);
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
        AppelEnCours:=True ;
        Fo_PBC.AfficheFeu;
        CtiModeAppel := ModeAppelSortant;
        FicheMontee:=true;
        StCti:=RTCtiMonteFicheContact('CTI;NUMEROTER',StClient,ValeurI(StNum));
        if StCti <> '' then
            Result:=ReadToKenPipe(StCti,'|')+';'+StNat+';'+StClient+';'+StNum+';'+ReadToKenPipe(StCti,'|');
        FicheMontee:=false;
        CtiRaccrocherFiche:=False;
        if (CtiResultatAppel <> 0) then
            begin
            AppelEnCours:=False ;
            AfficheFeu;
            end;
        end
    else
        begin
        V_PGI.ZoomOLE := true;
        Result:=AGLLanceFiche('YY','YYCONTACT','T;'+StClient,StNum,'CTI') ;
        V_PGI.ZoomOLE := false;
        end;
    end;
Until Result='';
end;
procedure TForm1.TiersClick(Sender: TObject);
begin
V_PGI.ZoomOLE := true;
AGLLanceFiche ('RT', 'RTCTI_RECHTIERS','' , '', '');
V_PGI.ZoomOLE := false;
end;
procedure TForm1.ActionsClick(Sender: TObject);
begin
V_PGI.ZoomOLE := true;
AGLLanceFiche ('RT', 'RTCTI_ACTIONS','' , '', '');
V_PGI.ZoomOLE := false;
end;

procedure TForm1.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

end.

