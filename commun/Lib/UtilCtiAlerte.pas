unit UtilCtiAlerte;
 
interface
uses  Utob
{$IFDEF EAGLCLIENT}
      ,MaineAGL
{$ELSE}
      ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main
{$ENDIF}
      ,EntRT, Sysutils, HCtrls, HEnt1, HMsgBox
      ,M3FP, ctiConst,Utom,ParamSoc,ctiInterface
      ;

procedure RTCreatActionCTI ( ctiCall:TCtiCall; ctiContact : Tob; connectEvent : boolean) ;
Procedure RTUpdateActionCTI (cticall : TCtiCall; ctiContact : Tob);

implementation

procedure RTCreatActionCTI ( ctiCall:TCtiCall; ctiContact : Tob; connectEvent : boolean) ;
var TA,TobActGen : Tob;
    TM : Tom;
    b_result:boolean;
    Q : TQuery;
    NumAct : integer;
begin
  if (ctiCall = nil) or (ctiContact = nil) then exit;
  { SO_RTCTIGENACTEFFECT : appel sortant abouti
    SO_RTCTIGENACTNONEFF : appel sortant non abouti
    SO_RTCTIGENACTENTOK  : appel entrant abouti
    SO_RTCTIGENACTENTNOK : appel entrant non abouti
    SO_RTCTIMODACTEFFECT :
  }
  { On test si le type d'appel est à générer en action }

  { appel entrant abouti }
  if ( (ctiCall.Origin = ctiOrigin_Incoming) and (connectEvent)
        and (not GetParamsocSecur('SO_RTCTIGENACTENTOK',false)) ) or
     { appel Sortant abouti }
     ( (ctiCall.Origin = ctiOrigin_Outgoing) and (connectEvent)
        and (not GetParamsocSecur('SO_RTCTIGENACTEFFECT',false)) ) or
     { appel entrant non abouti }
     ( (ctiCall.Origin = ctiOrigin_Incoming) and (not connectEvent)
        and (not GetParamsocSecur('SO_RTCTIGENACTENTNOK',false)) ) or
     { appel Sortant non abouti }
     ( (ctiCall.Origin = ctiOrigin_Outgoing) and (not connectEvent)
        and (not GetParamsocSecur('SO_RTCTIGENACTNONEFF',false)) )
    then exit;

  VH_RT.RTModActionsCTI.Load(true);

  TM := CreateTom('ACTIONS', nil, False, True);
  TA:=Tob.create('ACTIONS',nil,-1);
  TM.InitTob(TA);
  { aboutis }
  if connectEvent then
    begin
    if ctiCall.Origin = ctiOrigin_Incoming then
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
                [GetParamsocSecur('SO_RTCTIMODENTOK','0')], false)
    else
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTITYPEEFFECTUES','0')], false);
    end
  else
    { non aboutis }
    begin
    if ctiCall.Origin = ctiOrigin_Incoming then
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTIMODENTNOK','0')], false)
    else
       TobActGen:=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTITYPENONEFFECT','0')], false);
    end;

  if not Assigned(TobActGen) then
      BEGIN
      PGIBox('Modèle d''action non trouvé', 'Appel CTI');
      if TM <> NIL then TM.Free ;
      if TA <> NIL then TA.Free ;
      exit;
      end;

  NumAct:=1;
  Q:=OpenSQL('SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+Tob(ctiContact).GetString('T_AUXILIAIRE')+'"', True);
  if not Q.Eof then NumAct := Q.Fields[0].AsInteger+1;
  Ferme(Q) ;
  TA.PutValue ('RAC_TIERS',Tob(ctiContact).GetString('T_TIERS'));
  TA.PutValue ('RAC_AUXILIAIRE',Tob(ctiContact).GetString('T_AUXILIAIRE'));
  TA.PutValue ('RAC_LIBELLE',TobActGen.GetValue('RAG_LIBELLE'));
  TA.PutValue ('RAC_TYPEACTION',TobActGen.GetValue('RAG_TYPEACTION'));

  TA.PutValue ('RAC_DATEECHEANCE',TobActGen.GetValue('RAG_DATEECHEANCE'));
  TA.PutValue ('RAC_ETATACTION',TobActGen.GetValue('RAG_ETATACTION'));
  TA.PutValue ('RAC_TABLELIBRE1',TobActGen.GetValue('RAG_TABLELIBRE1'));
  TA.PutValue ('RAC_TABLELIBRE2',TobActGen.GetValue('RAG_TABLELIBRE2'));
  TA.PutValue ('RAC_TABLELIBRE3',TobActGen.GetValue('RAG_TABLELIBRE3'));
  TA.PutValue ('RAC_COUTACTION',TobActGen.GetValue('RAG_COUTACTION'));

  TA.putvalue('RAC_NUMACTION',NumAct );
  TA.SetDatetime('RAC_HEUREACTION', cticall.TimeStart);
  TA.putvalue('RAC_INTERVENANT', VH_RT.RTResponsable);
  TA.putvalue('RAC_DATECREATION', Date);
  if Tob(ctiContact).GetInteger('C_NUMEROCONTACT') <> 0 then
     TA.putvalue('RAC_NUMEROCONTACT', Tob(ctiContact).GetInteger('C_NUMEROCONTACT'));
  if (TM.VerifTOB( TA )) then
    begin
    b_result := TA.InsertDB (nil);
    If Not b_result Then PGIBox('Impossible de générer l''action', 'Appel CTI')
    end else
        PGIBox(TM.LastErrorMsg, 'Appel CTI');

  if TM <> NIL then TM.Free ;
  if TA <> NIL then  TA.Free ;

  { mng no action cticall }
  // $$$ JP 13/08/07
  ctiContact.PutValue ('RAC_NUMACTION', NumAct);
  //ctiCall.NumAction := NumAct;
end;

Procedure RTUpdateActionCTI (cticall : TCtiCall; ctiContact : Tob);
var Chrono : double;
    TobActGen : tob;
    Auxi,Sql : String;
begin
   if (ctiCall = nil) or (ctiContact = nil) then
      exit;
   Auxi := ctiContact.GetString ('T_AUXILIAIRE');

  Chrono := ctiCall.ConnectedTime;

  TobActGen := nil;
  if ctiCall.Origin = ctiOrigin_Outgoing then
    begin
    { appel sortant abouti }
    if GetParamsocSecur('SO_RTCTIGENACTEFFECT',false) then
      TobActGen :=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
          [GetParamsocSecur('SO_RTCTITYPEEFFECTUES','0')], false)
    end
  else
    { appel entrant abouti }
    begin
    if GetParamsocSecur('SO_RTCTIGENACTENTOK',false) then
      TobActGen :=VH_RT.RTModActionsCTI.FindFirst(['RAG_NUMACTGEN'],
              [GetParamsocSecur('SO_RTCTIMODENTOK','0')], false)
    end ;
  if Assigned(TobActGen) then
    begin
    Sql:='Update ACTIONS set ';
    if Chrono <> 0 then
      Sql:=Sql+'RAC_CHRONOMETRE='+StrFPoint(Chrono)+',';
    Sql:=Sql+'RAC_TYPEACTION="'+TobActGen.GetString('RAG_TYPEACTION')+'"'+ ',RAC_LIBELLE="'+TobActGen.GetString('RAG_LIBELLE')+'"'+
      ' Where RAC_AUXILIAIRE="'+Auxi+'" and RAC_NUMACTION=' + IntToStr (ctiContact.GetValue ('RAC_NUMACTION')); // $$$ JP 13/08/07 IntToStr(ctiCall.NumAction);
    ExecuteSQL (Sql);
    end;
end;

end.
