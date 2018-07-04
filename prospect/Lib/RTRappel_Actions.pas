unit RTRappel_Actions;

interface

uses Classes,
  windows,
{$IFDEF EAGLCLIENT}
     Maineagl, Utob,
{$ELSE}
     FE_Main,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Db,
{$ENDIF}
     sysutils,HMsgBox,Hent1
     ,M3Fp,EntRT,HCtrls,ParamSoc
     ;

procedure RTInitTHreadActions(ToutDeSuite : boolean);
procedure RTEndThreadActions();
Procedure AGLRTAfficheRappelsActions ( parms: array of variant; nb: integer );

type
  TThreadActions = class(TThread)
    procedure ThreadAfficheRappelsActions() ;
    constructor Create();
    destructor Destroy; OVERRIDE;

    procedure Execute; OVERRIDE;
  end;

implementation

var
  LeThreadActions: TThreadActions;
  JeSuisArrete: boolean ;
  Duree : Integer;
  Immediatement : boolean;

procedure AfficheRappelsActions ;
var Retour : String;
    requete : string;
    Q : TQuery;
    NbChamps : integer;
begin
    requete:='SELECT count(*)  FROM RTVACTIONS WHERE RAC_DATERAPPEL<>"'+UsDateTime(idate1900)+'" and RAC_ETATACTION="PRE"';
    requete:=requete+' AND RAC_GESTRAPPEL="X" AND RAC_DATERAPPEL <= "'+UsDateTime(Now);
    requete:=requete+'" AND RAC_INTERVENANT="'+VH_RT.RTResponsable+'"'+VH_RT.RTConfWhereConsult;
    Q:=OpenSQL(requete,True,-1,'',true);
    if not Q.EOF then
       NbChamps:=Q.Fields[0].AsInteger
    else NbChamps:=0;
    Ferme(Q) ;
    if NbChamps <> 0 then
    begin
        if V_PGI.ListeFiches.count <> 0 then
           if copy(V_PGI.ListeFiches[V_PGI.ListeFiches.count-1],1,17) = 'RTACTIONS_MUL_RAP' then exit;
        V_PGI.ZoomOLE := true;  //Affichage en mode modal
        Retour:=AGLLanceFiche('RT','RTACTIONS_MUL_RAP','','','RAPPEL') ;
        V_PGI.ZoomOLE := False;  //Affichage en mode modal
    end;
end;

procedure RTInitTHreadActions(ToutDeSuite : boolean);
var NbMin : String;
begin
  V_PGI.Debug:=False;
  NbMin:=GetSynRegKey('RTRappelMinutes','',true);
  if NbMin='' then
     Duree:=GetParamSocSecur('SO_RTACTDUREE',10)*60*1000
  else
     Duree:=StrToInt(NbMin)*60*1000;
  if ToutDeSuite then AfficheRappelsActions;
  Immediatement:=False;
  if not assigned(LeThreadActions) then LeThreadActions := TThreadActions.Create();
end;

procedure RTEndThreadActions();
begin
  if assigned(LeThreadActions) then
    begin
    LeThreadActions.terminate;
    LeThreadActions:=nil;
    end;
end;

Procedure AGLRTAfficheRappelsActions ( parms: array of variant; nb: integer );
begin
//  if V_PGI.MenuCourant = 92 then { mng 21/09/2004, si fenêtre du lancement, on ne fait pas }
//    begin
    SaveSynRegKey('RTRappelAuto',parms[0],TRUE);
    SaveSynRegKey('RTRappelMinutes',parms[1],TRUE);

    if (parms[0] = 'X') and  (not assigned(LeThreadActions)) then
      {démarrage}
      RTInitTHreadActions(False)
    else
      if (parms[0] = '-') and (assigned(LeThreadActions)) then
        {arrêt}
        RTEndThreadActions()
      else
        if (parms[0] = 'X') then
        {prise en compte nouvelle durée : suspension et reprise}
        begin
          LeThreadActions.Suspend;
          Duree:=StrToInt(parms[1])*60*1000;
          LeThreadActions.Resume;
        end;
//     end;
end;


{ TThreadActions }

procedure TThreadActions.ThreadAfficheRappelsActions;
begin
  AfficheRappelsActions;
end;


constructor TThreadActions.Create;
var IDuree : integer;
begin
  inherited Create(True);

  JeSuisArrete := false ;

  { Suppression à la fin de l'exécution }
  FreeOnTerminate := True;

  // Chargement des données de la table...
  if Immediatement then
  begin
    // Réveille toi doucement
    IDuree:=Duree;
    Duree:=2000;
    Resume;
    sleep(4000);
    Suspend;
    Duree:=IDuree;
  end;
  // Réveille toi pour de bon
  Resume;

end;

destructor TThreadActions.Destroy;
begin
  inherited;
  if (V_PGI.Debug = true) and (JeSuisArrete = False ) then
     PGIInfo('Attention : le rappel automatique est interrompu par le DebugLog','Rappels Actions');
  JeSuisArrete:=true ;
  LeThreadActions := Nil ;
end;

procedure TThreadActions.Execute;

  function Attente : boolean;
  var
    i, il: integer;
  begin
    i := 0;
    il := Duree;
    result := True;
    while (i < il) and result do
    begin
      if Terminated then result := False
      else
      begin
        i := i + 2000;
        sleep(2000);
        il := Duree;
      end;
    end;
  end;
begin
  while (not Terminated) do
  begin
    if Attente then
    begin
      // A toi ...
        // if true {and ExisteAction} then Synchronize(AfficheRappelsActions) ;
       Synchronize( ThreadAfficheRappelsActions);
    end else break ;
  end;
end;


Initialization
  RegisterAglProc( 'RTAfficheRappelsActions', False , 2, AGLRTAfficheRappelsActions);
  LeThreadActions := Nil ;
  JeSuisArrete := True ;

Finalization
  If assigned(LeThreadActions) then LeThreadActions.Terminate ;
  JeSuisArrete := True ;
end.


