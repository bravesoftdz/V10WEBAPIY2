{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPECHEMONO ()
Mots clefs ... : TOF;CPECHEMONO
*****************************************************************}
Unit CPECHEMONO_TOF ;

Interface

Uses Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEagl,
{$ELSE}
     Fe_Main,
{$ENDIF}
     uTob,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,             // TFVierge
     Saisutil,           // NbJoursOK
     UTOF ;


function CPLanceFiche_EcheMono( vStArg : String ) : boolean ;

Type
  TOF_CPECHEMONO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    // EVT Fiche
    procedure E_ModePaieChanged        ( Sender : TObject ) ;
    procedure FormKeyDown              ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
  private
    FTobEche      : TOB ;
    FModalResult  : TModalResult ;
  end ;

Implementation

uses Windows ;           // VK_ESCAPE

Const MessageListe : Array[1..3] of String =(
        {01}                    'Vous devez renseigner un mode de paiement.',
        {02}                    'Vous devez valider les informations.',
        {03}                    'La date d''échéance doit respecter la plage de saisie autorisée.'
		                            );

function CPLanceFiche_EcheMono( vStArg : String ) : boolean ;
begin
  if pos('ACTION=', vStArg) = 0 then
     vStArg := 'ACTION=CONSULTATION;' + vStArg ;
  result := AGLLanceFiche( 'CP', 'CPECHEMONO', '', '', vStArg ) = 'OK' ;
end ;


procedure TOF_CPECHEMONO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMONO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMONO.OnUpdate ;
var lDateEche : TDateTime ;
begin
  Inherited ;

  if ( GetControlText('E_MODEPAIE')='' ) and ( TFVierge(Ecran).TypeAction <> taConsult ) then
    begin
    PgiInfo( TraduireMemoire( MessageListe[1] ), Ecran.Caption ) ;
    LastError := 1 ;
    end ;

  if ( FTobEche.GetNumChamp('E_DATECOMPTABLE') > 0 ) and
     ( FTobEche.GetValue('E_DATECOMPTABLE') <> iDate1900 ) then
     begin
     lDateEche := StrToDate( GetControlText('E_DATEECHEANCE') ) ;
     if Not NbJoursOK( FTobEche.GetValue('E_DATECOMPTABLE'), lDateEche ) then
       begin
       PgiInfo( TraduireMemoire( MessageListe[3] ), Ecran.Caption ) ;
       LastError := 1 ;
       end ;
     end ;
{
  if Not IsValidDate( FDateEche.Text ) then
    FDateEche.Text := DateToStr( EcheInit ) ;
  if Not IsValidDate( FDateValeur.Text ) then
    FDateValeur.Text := DateToStr( ValInit ) ;
}

  FTobEche.GetEcran( Ecran ) ;
  FModalResult := mrOk ;
  TFVierge(Ecran).laTof.laTob := FTobEche ;

end ;

procedure TOF_CPECHEMONO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMONO.OnArgument (S : String ) ;
begin
  Inherited ;

  THValComboBox( Getcontrol('E_MODEPAIE', True)).OnChange := E_ModePaieChanged ;
  TFVierge(Ecran).OnKeyDown                               := FormKeyDown ;

  FTobEche := TFVierge(Ecran).laTof.laTob ;
  FTobEche.PutEcran( Ecran ) ;

  // Variables initalisées par les paramètres d'appels...
  FModalResult         := mrNone ;

end ;

procedure TOF_CPECHEMONO.OnClose ;
begin
  Inherited ;
  if FModalResult = mrOk
    then TFVierge(Ecran).retour := 'OK'
    else TFVierge(Ecran).retour := 'CANCEL' ;
end ;

procedure TOF_CPECHEMONO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMONO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPECHEMONO.E_ModePaieChanged(Sender: TObject);
var lStCat : String ;
begin
  lStCat := MPToCategorie( GetControlText('E_MODEPAIE') ) ;

  SetControlEnabled( 'E_NUMTRAITECHQ',  (lStCat='CHQ') or (lStCat='LCR') ) ;
  SetControlEnabled( 'TE_NUMTRAITECHQ', (lStCat='CHQ') or (lStCat='LCR') ) ;

  if lStCat = 'CHQ'
    then SetControlText( 'TE_NUMTRAITECHQ', TraduireMemoire('N° Chèque') )
    else SetControlText( 'TE_NUMTRAITECHQ', TraduireMemoire('N° Traite') ) ;

end;

procedure TOF_CPECHEMONO.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ( Key = VK_ESCAPE )  then
    FModalResult := mrCancel ;
  TFVierge(Ecran).FormKeyDown( Sender, Key, Shift ) ;
end;

Initialization
  registerclasses ( [ TOF_CPECHEMONO ] ) ;
end.
