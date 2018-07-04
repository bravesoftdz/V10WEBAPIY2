unit UtilCti;

interface
uses  HCtrls,HEnt1,UtomAction,UtofRTCourrier,EntRT,ParamSoc,Utob,
      Utom ,HMsgBox, SysUtils
{$IFDEF EAGLCLIENT}
      , Maineagl
{$ELSE}
      , FE_main {$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF}
      , UtilPgi  
;
Const ModeleAction : String = 'MODELES D''ACTIONS';
      ModeAppelSortant : integer = 1 ;
      ModeAppelEntrant : integer = 2 ;

Function  RTCtiAfficheClient ( appelant : WideString; TypeMontee : String ) : String;
Function  RTCtiMonteFicheClient ( TypeMontee, StClient, StNat : String ) : String;
function RTCTIGenereAction (AppelCtiOk : Boolean; StTiers,StAuxi : String; HeureDeb,Duree : TDateTime; WhereUpdate : String; ModeAppel,NumeroContact : integer) : Integer;
procedure RTCTIMajDureeAction (StAuxi : String; NumAction  : integer; Duree : TDateTime );
Function  RTCtiAfficheContact ( appelant : WideString; TypeMontee : String ) : String;
Function  RTCtiMonteFicheContact ( TypeMontee, StClient : String; INum : integer ) : String;
implementation

Function  RTCtiAfficheClient ( appelant : WideString; TypeMontee : String ) : String;
var    QQ :tquery;
       StClient,StNat,StCti,StLib,StTiers,StNoTel : String;
begin
StLib:=''; StTiers:='';
Result:=''; StCti:=''; StNat:='';
V_PGI.ZoomOLE := false;
StNotel:=appelant;
if GetParamsocSecur('SO_RTSUPPZERO',false) then
  StNotel:='0'+appelant;
QQ := OpenSQL('Select T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_TIERS from TIERS Where T_CLETELEPHONE="'+CleTelephone(StNotel)+'"',True) ;
if Not QQ.EOF then
  begin
  StClient:=QQ.FindField('T_AUXILIAIRE').AsString;
  StNat:=QQ.FindField('T_NATUREAUXI').AsString;
  StLib:=QQ.FindField('T_LIBELLE').AsString;
  StTiers:=QQ.FindField('T_TIERS').AsString;
  Ferme(QQ);
  StCti:=RTCtiMonteFicheClient(TypeMontee,StClient,StNat) ;
  end
else
  Ferme(QQ);
Result:=StTiers+';'+StLib+';'+StNat+';'+StClient;
end;

Function  RTCtiMonteFicheClient ( TypeMontee, StClient, StNat : String ) : String;
var StCti :String;
begin
if StClient <> '' then
    begin
    V_PGI.ZoomOLE := true;  //Affichage en mode modal
    StCti:=AGLLanceFiche('GC','GCTIERS','',StClient,TypeMontee+';T_NATUREAUXI='+StNat) ;
    V_PGI.ZoomOLE := false;
    Result:=StCti;  // Code tiers+libelle
    end;
end;

Function  RTCtiAfficheContact ( appelant : WideString; TypeMontee : String ) : String;
var    QQ :tquery;
       StClient,StNat,StCti,StLib,StTiers : String;
       INum : integer;
begin
StLib:=''; StTiers:='';
Result:=''; StCti:=''; StNat:='';
V_PGI.ZoomOLE := false;
INum:=0;
QQ := OpenSQL('Select C_NOM,C_AUXILIAIRE,C_NATUREAUXI,C_NUMEROCONTACT from CONTACT Where C_CLETELEPHONE="'+CleTelephone(appelant)+'" and C_TYPECONTACT="T"',True) ;
if Not QQ.EOF then
  begin
  StClient:=QQ.FindField('C_AUXILIAIRE').AsString;
  StNat:=QQ.FindField('C_NATUREAUXI').AsString;
  //StLib:=QQ.FindField('T_LIBELLE').AsString;
  //StTiers:=QQ.FindField('T_TIERS').AsString;
  INum:=QQ.FindField('C_NUMEROCONTACT').AsInteger;
  Ferme(QQ);
  StCti:=RTCtiMonteFicheContact(TypeMontee,StClient,INum) ;
  end
else
  Ferme(QQ);
Result:=ReadToKenPipe(StCti,'|')+';'+StNat+';'+StClient+';'+IntToStr(INum)+';'+ReadToKenPipe(StCti,'|');
end;

Function  RTCtiMonteFicheContact ( TypeMontee, StClient : String; INum : integer ) : String;
var StCti :String;
begin
if StClient <> '' then
    begin
    V_PGI.ZoomOLE := true;  //Affichage en mode modal
    StCti:=AGLLanceFiche('YY','YYCONTACT','T;'+StClient,IntToStr(INum),TypeMontee) ;
    V_PGI.ZoomOLE := false;
    Result:=StCti;  // Code tiers+libelle
    end;
end;


function RTCTIGenereAction (AppelCtiOk : Boolean; StTiers,StAuxi : String; HeureDeb,Duree : TDateTime; WhereUpdate : String; ModeAppel,NumeroContact : integer) : Integer;
var TA,TobActGen : Tob;
    TM : Tom_actions;
    b_result:boolean;
    requete : string;
    NumAct : integer;
begin
Result:=0;
if ( (AppelCtiOk) and (not GetParamsocSecur('SO_RTCTIGENACTEFFECT',false)) ) or
   ( (not AppelCtiOk) and (not GetParamsocSecur('SO_RTCTIGENACTNONEFF',false)) ) then exit;


if (AppelCtiOk) and (GetParamsocSecur('SO_RTCTIMODACTEFFECT',false))
    and ( GetParamsocSecur('SO_RTCTITYPEEFFECTUES','') <> '') and (WhereUpdate<>'') then
    begin
    requete:='Update actions set rac_etataction=(Select rag_etataction from '+
    'ACTIONSGENERIQUES Where RAG_OPERATION="'+ModeleAction+'" AND RAG_NUMACTGEN ='
    + GetParamsocSecur('SO_RTCTITYPEEFFECTUES','')+')'+
    ',RAC_DATEMODIF="'+USDateTime(Date)+'"'+
    ',RAC_HEUREACTION="'+USDateTime(HeureDeb)+'"'+
    ',RAC_DATEACTION="'+USDateTime(Date)+'" where rac_auxiliaire="'+StAuxi+'"'+WhereUpdate;
    ExecuteSQL(requete);
    exit;
    end;

VH_RT.RTModActionsCTI.Load(true);

TM:=Tom_Actions(CreateTOM('ACTIONS',Nil,False,False) );
TM.stTiers:=StTiers ;
TA:=Tob.create('ACTIONS',nil,-1);
TM.InitTob(TA);
if AppelCtiOk then
    begin
    if ModeAppel = ModeAppelSortant then
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTITYPEEFFECTUES','0')], false)
    else
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
                [GetParamsocSecur('SO_RTCTIMODENTOK','0')], false);
    end
else
    begin
    HeureDeb:=Time; // non renseigné si pas eu d'appel
    if ModeAppel = ModeAppelSortant then
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTITYPENONEFFECT','0')], false)
    else
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTIMODENTNOK','0')], false);             
    end;

  if not Assigned(TobActGen) then
    BEGIN
    PGIBox('Modèle d''action non trouvé', 'Appel CTI');
    if TM <> NIL then TM.Free ;
    if TA <> NIL then TA.Free ;
    exit;
    end;

TA.PutValue ('RAC_LIBELLE',TobActGen.GetValue('RAG_LIBELLE'));
TA.PutValue ('RAC_TYPEACTION',TobActGen.GetValue('RAG_TYPEACTION'));

TA.PutValue ('RAC_DATEECHEANCE',TobActGen.GetValue('RAG_DATEECHEANCE'));
TA.PutValue ('RAC_ETATACTION',TobActGen.GetValue('RAG_ETATACTION'));
TA.PutValue ('RAC_TABLELIBRE1',TobActGen.GetValue('RAG_TABLELIBRE1'));
TA.PutValue ('RAC_TABLELIBRE2',TobActGen.GetValue('RAG_TABLELIBRE2'));
TA.PutValue ('RAC_TABLELIBRE3',TobActGen.GetValue('RAG_TABLELIBRE3'));
TA.PutValue ('RAC_COUTACTION',TobActGen.GetValue('RAG_COUTACTION'));

NumAct:=GetNextAction(StAuxi);
TA.putvalue('RAC_NUMACTION',NumAct );
TA.putvalue('RAC_HEUREACTION', HeureDeb);
if AppelCtiOk then
    TA.putvalue('RAC_CHRONOMETRE', Duree);
//if GetParamSoc('SO_RTACTRESP')='COM'  then
//   TA.putvalue('RAC_INTERVENANT', RTRechResponsable (stTiers));
// plus logique de prendre le code responsable associé à l'utilisateur
TA.putvalue('RAC_INTERVENANT', VH_RT.RTResponsable);
TA.putvalue('RAC_DATECREATION', Date);
if NumeroContact <> 0 then
   TA.putvalue('RAC_NUMEROCONTACT', NumeroContact);
if (TM.VerifTOB( TA )) then
  begin
  b_result := TA.InsertDB (nil);
  If Not b_result Then PGIBox('Impossible de générer l''action', 'Appel CTI')
  end else
      PGIBox(TM.LastErrorMsg, 'Appel CTI');

TobActGen.free;
if TM <> NIL then TM.Free ;
if TA <> NIL then  TA.Free ;
Result:= NumAct;
end;

procedure RTCTIMajDureeAction (StAuxi : String; NumAction  : integer; Duree : TDateTime );
var Chrono : double;
begin
Chrono:=Double(Duree);
ExecuteSQL ('Update ACTIONS set RAC_CHRONOMETRE='+StrFPoint(Chrono)+' Where RAC_AUXILIAIRE="'+StAuxi+
     '" and RAC_NUMACTION='+IntToStr(NumAction));
end;
end.
