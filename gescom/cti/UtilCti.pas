unit UtilCti;

interface
uses  HCtrls,HEnt1,UtomAction,UtofRTCourrier,EntRT,ParamSoc,Utob,
      Utom ,HMsgBox, SysUtils
{$IFDEF EAGLCLIENT}
      , Maineagl
{$ELSE}
      , FE_main {$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF}
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
var	QQ 				: TQuery;
		Req				: String;
    StTiers 	: string;
    StAux     : String;
    StNat			: string;
    StCti			: string;
    StNoTel 	: String;
    StLContact: String;
    StLibTiers: String;
    StAdr1		: String;
    StAdr2		: String;
    StAdr3		: String;
    StCP			: String;
    StVille		: String;
begin

  Result    :='';

  StTiers  	:= '';
  StAux     := '';
  StNat			:= '';
  StCti			:= '';
  StLContact:= '';
  StLibTiers:= '';
  StAdr1		:= '';
  StAdr2		:= '';
  StAdr3		:= '';
  StCP			:= '';
  StVille		:= '';
  StNotel   := Appelant;

	V_PGI.ZoomOLE := false;

  Result := StNoTel;

	if GetParamSoc('SO_RTSUPPZERO') then StNotel := '0'+appelant;

	//Recherche du contact en fonction du N° de Téléphone
  Req := 'Select C_TIERS, C_NATUREAUXI, C_NOM ';
  Req := Req + 'From CONTACT Where C_CLETELEPHONE ="';
  Req := Req + CleTelephone(StNotel) + '"';

  QQ := OpenSQL(Req, True);
  If Not QQ.EOF Then
  	 Begin
     StLContact := QQ.FindField('C_NOM').AsString;
     StTiers := QQ.FindField('C_TIERS').AsString;
	   StNat:=QQ.FindField('C_NATUREAUXI').AsString;
     Req := 'Select T_AUXILIAIRE, T_TIERS, T_NATUREAUXI, ';
     Req := Req + 'T_LIBELLE, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, ';
     Req := Req + 'T_CODEPOSTAL, T_VILLE From TIERS Where T_TIERS="';
     Req := Req + StTiers + '" And T_NATUREAUXI="' + StNat + '"';
     end
	Else
  	 Begin
     Req := 'Select T_AUXILIAIRE, T_TIERS,T_NATUREAUXI,T_LIBELLE, ';
     Req := Req + 'T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, ';
     Req := Req + 'T_CODEPOSTAL, T_VILLE ';
     Req := Req + 'From TIERS Where T_CLETELEPHONE="';
     Req := Req + CleTelephone(StNotel) + '"';
     end;

  //Recherche du Tiers
  QQ := OpenSQL(Req,True);

	if Not QQ.EOF then
	   begin
  	 StTiers:=QQ.FindField('T_TIERS').AsString;
  	 StAux:=QQ.FindField('T_TIERS').AsString;
	   StNat:=QQ.FindField('T_NATUREAUXI').AsString;
	   StLibTiers:=QQ.FindField('T_LIBELLE').AsString;
     StAdr1 := QQ.FindField('T_ADRESSE1').AsString;
     StAdr2 := QQ.FindField('T_ADRESSE2').AsString;
     StAdr3 := QQ.FindField('T_ADRESSE3').AsString;
		 StCP := QQ.FindField('T_CODEPOSTAL').AsString;
     StVille := QQ.FindField('T_VILLE').AsString;
	   end
  else
     Begin
     Req := 'Select ADR_CONTACT, ADR_REFCODE, ADR_NATUREAUXI,'+
            'ADR_ADRESSE1, ADR_ADRESSE2, ADR_ADRESSE3,'+
            'ADR_CODEPOSTAL, ADR_VILLE, ADR_LIBELLE '+
            'From ADRESSES '+
            'LEFT JOIN BADRESSES ON BA0_TYPEADRESSE=ADR_TYPEADRESSE AND BA0_NUMEROADRESSE=ADR_NUMEROADRESSE '+
            'Where BA0_CLETELEPHONE="'+ CleTelephone(StNotel) + '" And BA0_INT="X"';
	   //Recherche de l'adresse d'intervention en fonction de N° Téléphone
  	 QQ := OpenSQL(Req,True);
	   if Not QQ.EOF then
  	  	begin
  	 		StTiers:=QQ.FindField('ADR_REFCODE').AsString;
        StAux:=QQ.FindField('ADR_REFCODE').AsString;
        StNat:=QQ.FindField('ADR_NATUREAUXI').AsString;
        StLibTiers:=QQ.FindField('ADR_LIBELLE').AsString;
        StAdr1 := QQ.FindField('ADR_ADRESSE1').AsString;
        StAdr2 := QQ.FindField('ADR_ADRESSE2').AsString;
        StAdr3 := QQ.FindField('ADR_ADRESSE3').AsString;
        StCP := QQ.FindField('ADR_CODEPOSTAL').AsString;
        StVille := QQ.FindField('ADR_VILLE').AsString;
        end;
     end;

  Result := Result + ';' + StLContact + ';' + StAux;
  Result := Result + ';' + StTiers + ';' + StLibTiers;
  Result := Result + ';' + StNat + ';' + StAdr1 + ';' + StAdr2;
  Result := Result + ';' + StAdr3 + ';' + StCP + ';' + StVille;

	if result = Appelant then Result := '';

  Ferme(QQ);

end;

Function  RTCtiMonteFicheClient ( TypeMontee, StClient, StNat : String ) : String;
var StCti :String;
begin
if StClient <> '' then
    begin
    V_PGI.ZoomOLE := true;  //Affichage en mode modal
    //StCti:=AGLLanceFiche('GC','GCTIERS','',StClient,TypeMontee+';T_NATUREAUXI='+StNat) ;
    StCti:=AGLLanceFiche('BTP','BTAPPELINT','','','ACTION=CREATION;NUMTEL=' + StClient) ;
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
    QAG : TQuery;
    NumAct : integer;
begin
Result:=0;
if ( (AppelCtiOk) and (not GetParamSoc('SO_RTCTIGENACTEFFECT')) ) or
   ( (not AppelCtiOk) and (not GetParamSoc('SO_RTCTIGENACTNONEFF')) ) then exit;


if (AppelCtiOk) and (GetParamSoc('SO_RTCTIMODACTEFFECT')) and (WhereUpdate<>'') then
    begin
    requete:='Update actions set rac_etataction=(Select rag_etataction from '+
    'ACTIONSGENERIQUES Where RAG_OPERATION="'+ModeleAction+'" AND RAG_NUMACTGEN ='
    + GetParamSoc('SO_RTCTITYPEEFFECTUES')+')'+
    ',RAC_DATEMODIF="'+USDateTime(Date)+'"'+
    ',RAC_HEUREACTION="'+USDateTime(HeureDeb)+'"'+
    ',RAC_DATEACTION="'+USDateTime(Date)+'" where rac_auxiliaire="'+StAuxi+'"'+WhereUpdate;
    ExecuteSQL(requete);
    exit;
    end;


TM:=Tom_Actions(CreateTOM('ACTIONS',Nil,False,False) );
TM.stTiers:=StTiers ;
TA:=Tob.create('ACTIONS',nil,-1);
TM.InitTob(TA);
if AppelCtiOk then
    begin
    if ModeAppel = ModeAppelSortant then
        QAG := OpenSQL('Select * from ACTIONSGENERIQUES Where RAG_OPERATION="' +
             ModeleAction + '" AND RAG_NUMACTGEN ='+ GetParamSoc('SO_RTCTITYPEEFFECTUES'),True)
    else
        QAG := OpenSQL('Select * from ACTIONSGENERIQUES Where RAG_OPERATION="' +
             ModeleAction + '" AND RAG_NUMACTGEN ='+ GetParamSoc('SO_RTCTIMODENTOK'),True);
    end
else
    begin
    HeureDeb:=Time; // non renseigné si pas eu d'appel
    if ModeAppel = ModeAppelSortant then
        QAG := OpenSQL('Select * from ACTIONSGENERIQUES Where RAG_OPERATION="' +
             ModeleAction + '" AND RAG_NUMACTGEN ='+ GetParamSoc('SO_RTCTITYPENONEFFECT'),True)
    else
        QAG := OpenSQL('Select * from ACTIONSGENERIQUES Where RAG_OPERATION="' +
             ModeleAction + '" AND RAG_NUMACTGEN ='+ GetParamSoc('SO_RTCTIMODENTNOK'),True);
    end;

if QAG.EOF then
    BEGIN
    PGIBox('Modèle d''action non trouvé', 'Appel CTI');
    if TM <> NIL then TM.Free ;
    if TA <> NIL then TA.Free ;
    exit;
    end
else
    begin
    TobActGen:=TOB.create ('ACTIONSGENERIQUES',NIL,-1);
    TobActGen.SelectDB('',QAG);
    end;
Ferme(QAG) ;
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
