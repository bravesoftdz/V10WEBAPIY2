{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPPARAMVIEWER ()
Mots clefs ... : TOF;CPPARAMVIEWER
*****************************************************************}
Unit CPPARAMVIEWER_TOF ;

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
     Vierge,               // TFVierge
//     Saisutil,           // NbJoursOK
     UTOF ;


function CPLanceFiche_ParamViewer( vStArg : String ) : boolean ;

Type
  TOF_CPPARAMVIEWER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    // EVT Fiche
    procedure FormKeyDown              ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
    procedure FormCloseQuery           ( Sender : TObject ; var CanClose: Boolean);

  private

   FBoActif      : boolean ;
   FStExt        : string ;
   FStPosition   : string ;
   FStChemin     : string ;
   FBoImportGED  : boolean ;
   
   FModalResult  : TModalResult ;

  end ;

Implementation

uses Windows,
     CBPPath,
     ParamSoc ;           // VK_ESCAPE

function CPLanceFiche_ParamViewer( vStArg : String ) : boolean ;
begin
  if pos('ACTION=', vStArg) = 0 then
     vStArg := 'ACTION=MODIFICATION;' + vStArg ;
  result := AGLLanceFiche( 'CP', 'CPPARAMVIEWER', '', '', vStArg ) = 'OK' ;
end ;


procedure TOF_CPPARAMVIEWER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMVIEWER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMVIEWER.OnUpdate ;
begin
  Inherited ;
  LastError := 0 ;

  // les tests ne sont valables que si on active le visualiseur en saisie
  if GetControlText('SO_CPAVECVIEWER') = 'X' then
    begin

    // définition de la position
    if ( GetControlText('SO_CPVIEWERPOS')='' ) then
      begin
      PgiInfo( TraduireMemoire( 'Vous devez renseignez la position du visualiseur' ), Ecran.Caption ) ;
      LastError := 1 ;
      Exit ;
      end ;

    // définition du répertoire
    if ( GetControlText('SO_CPVIEWERREP')='' ) or not DirectoryExists( GetControlText('SO_CPVIEWERREP') ) then
      begin
      PgiInfo( TraduireMemoire( 'Vous devez renseignez le répertoire des fichiers à traiter.' ), Ecran.Caption ) ;
      LastError := 1 ;
      Exit ;
      end ;

    // définition de l'extension de fichier
    if GetControlText('SO_CPVIEWEREXT')='' then
      begin
      PgiInfo( TraduireMemoire( 'Vous devez renseignez l''extension des fichiers à visualiser.' ), Ecran.Caption ) ;
      LastError := 1 ;
      Exit ;
      end ;

    end ;

  // Enregistrement des données
  SetParamSoc( 'SO_CPAVECVIEWER',      GetControlText('SO_CPAVECVIEWER')='X' ) ;
  SetParamSoc( 'SO_CPVIEWERPOS',       GetControlText('SO_CPVIEWERPOS') ) ;
  SetParamSoc( 'SO_CPVIEWERREP',       GetControlText('SO_CPVIEWERREP') ) ;
  SetParamSoc( 'SO_CPVIEWEREXT',       GetControlText('SO_CPVIEWEREXT') ) ;
  SetParamSoc( 'SO_CPVIEWERIMPORTGED', GetControlText('SO_CPVIEWERIMPORTGED')='X' ) ;

  // définition du répertoire
  FModalResult := mrOk ;

end ;

procedure TOF_CPPARAMVIEWER.OnLoad ;
begin
  Inherited ;

  // Mise en place des infos
  SetcontrolText( 'SO_CPAVECVIEWER',       BoolToStr_( FBoActif )  );
  SetControlText( 'SO_CPVIEWERPOS',        FStPosition  ) ;
  SetControlText( 'SO_CPVIEWERREP',        FStChemin    ) ;
  SetControlText( 'SO_CPVIEWEREXT',        FStExt       ) ;
  SetControlText( 'SO_CPVIEWERIMPORTGED',  BoolToStr_( FBoimportGed ) ) ;

end ;

procedure TOF_CPPARAMVIEWER.OnArgument (S : String ) ;
var lStRep : String ;
begin
  Inherited ;

  TFVierge(Ecran).OnKeyDown      := FormKeyDown ;
  TFVierge(Ecran).OnCloseQuery   := FormCloseQuery ;

  // Variables initalisées par les paramètres d'appels...
  FModalResult         := mrNone ;

  lStRep := TCBPPath.GetCegidUserDocument ;

  // Chargement des param soc
  FBoActif          := GetParamSocSecur( 'SO_CPAVECVIEWER', False, True ) ;
  FStExt            := GetParamSocSecur( 'SO_CPVIEWEREXT', 'PDF', True ) ;
  FStPosition       := GetParamSocSecur( 'SO_CPVIEWERPOS', '1BA', True) ;
  FStChemin         := GetParamSocSecur( 'SO_CPVIEWERREP', lStRep, True ) ;
  FBoImportGED      := GetParamSocSecur( 'SO_CPVIEWERIMPORTGED', False, True ) ;

end ;

procedure TOF_CPPARAMVIEWER.OnClose ;
begin
  Inherited ;
  if FModalResult = mrOk
    then TFVierge(Ecran).retour := 'OK'
    else TFVierge(Ecran).retour := 'CANCEL' ;
end ;

procedure TOF_CPPARAMVIEWER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPPARAMVIEWER.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_CPPARAMVIEWER.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ( Key = VK_ESCAPE )  then
    begin
    LastError := 0 ;
    FModalResult := mrCancel ;
    end ;
  TFVierge(Ecran).FormKeyDown( Sender, Key, Shift ) ;
end;

procedure TOF_CPPARAMVIEWER.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := LastError <> 1 ;
  if CanClose then
    TFVierge(Ecran).FormCloseQuery( Sender, CanClose ) ;
end;

Initialization
  registerclasses ( [ TOF_CPPARAMVIEWER ] ) ;
end.
