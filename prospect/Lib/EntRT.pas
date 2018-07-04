unit EntRT;

interface

uses Hctrls,UTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DB,
{$ENDIF}
{$ifdef GIGI}
        ent1,
{$endif}
     EntPgi,
     HEnt1,
     {$IFNDEF EAGLSERVER}
       {$IFNDEF ERADIO}
        {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
         ctiAlerte,
        {$ENDIF GIGI}
       {$ENDIF !ERADIO}
     {$ENDIF !EAGLSERVER}
     Paramsoc,sysUtils;

{Const
    NbFichiers: Integer = 4;
}
Type LaVariableRT = class{$IFDEF EAGLSERVER}(TLaVariable){$ENDIF EAGLSERVER}
     Private
     Public
     TobTypesAction, TobTypesChainage, TobChampsPro : TMemoryTob;
     TobChampsProMul,TobParamCl : TOB ;
{$ifdef GIGI}   //MCD 04/07/2005
     TobChampsDpMul: tob;
{$ENDIF GIGI}
     TobParamPlanning,TobParamEtatPlan : TMemoryTob;
     RTResponsable :string;
     RTNomResponsable :string;
     RTCodeCommercial :string;
     RTNomCommercial :string;
     RTConfWhereConsult : string ;
     RTConfWhereModif : string ;
     DroitModifTiers : boolean ;
     RTExisteConfident : boolean ;
     RFConfWhereConsult : string ;
     RFConfWhereModif : string ;
     RFDroitModifTiers : boolean ;
     RFExisteConfident : boolean ;
     RTCreatPropositions : boolean ;
     RTCreatActions : boolean ;
     RTCreatInfos : boolean;
     RTCreatContacts : boolean ;
     RFCreatInfos : boolean;
     RFCreatContacts : boolean ;
     RFCreatActions : boolean ;
     CTISeria : boolean;
     FlyDocSeria : boolean;
     {$IFNDEF EAGLSERVER}
       {$IFNDEF ERADIO}
        {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
         ctiAlerte : TFCtiAlerte;
        {$ENDIF GIGI}
       {$ENDIF !ERADIO}
     {$ENDIF !EAGLSERVER}
     { CTI : je memorise les types d'actions a creer }
     RTModActionsCTI: tMemoryTob;
     AncienModeCTI : boolean;
     Published
     end ;

{$IFNDEF EAGLSERVER}
Var VH_RT : LaVariableRT ;
{$ENDIF !EAGLSERVER}

{$IFDEF EAGLSERVER}
function VH_RT: LaVariableRT;
{$ENDIF EAGLSERVER}

Procedure InitLaVariableRT ;
Procedure ChargeParamsRT ;
Procedure LibereLaVariableRT ;
Function RTGetResponsable : string;
Procedure ChargeTobRT ;
{$ifdef GIGI}
Procedure ChargeTobDP (Ismul: boolean=true);
{$ENDIF GIGI}
Procedure ChargeTobRTFille ( TobParam : Tob );
Procedure ChargeTypesActions ;
Procedure ChargeTypesChainages ;
Procedure RTGetCommercial ;
Procedure RTGetConfidentialite ;
Function OnChangeUserGRC : boolean;
Procedure ChargeParamPlanning ;
Procedure ChargeParamEtatPlan ;
Const ModeleAction : String = 'MODELES D''ACTIONS';
implementation

function GetSqlRPP : string;
begin
  Result := 'Select * from RTPARAMPLANNING WHERE RPP_INTERVENANT ="'+V_PGI.User+'"';
end;

function GetSqlREP : string;
begin
  Result := 'Select * from RTETATPLANNING WHERE REP_INTERVENANT ="'+V_PGI.User+'"'
end;

function GetSqlRPA : string;
begin
  Result := 'Select * from PARACTIONS order by rpa_chainage,rpa_numligne';
end;

function GetSqlRPG : string;
begin
  Result := 'Select * from PARCHAINAGES';
end;

Function OnChangeUserGRC : boolean;
var Resp : String;
Begin
Resp:=RTGetResponsable ;
VH_RT.RTResponsable:=ReadToKenSt(Resp);
VH_RT.RTNomResponsable:=ReadToKenSt(Resp);
RTGetConfidentialite ;
RTGetCommercial;
ChargeParamPlanning;
ChargeParamEtatPlan;
Result:=True;
End;

Procedure RTGetConfidentialite ;
var    QQ :tquery;
       TobAcces,TobExist : TOB;
       i : integer ;
       Consult1,Consult2,Modif1,Modif2 : string;
       bConsult1,bConsult2,bModif1,bModif2 : boolean;
begin
  VH_RT.RTConfWhereConsult:='';
  VH_RT.RTConfWhereModif:='' ; Consult1:='';Consult2:='';Modif1:='';Modif2:='';
  VH_RT.DroitModifTiers := False;
  VH_RT.RTExisteConfident := False;
  VH_RT.RTCreatPropositions := False;
  VH_RT.RTCreatActions := False;
  VH_RT.RTCreatContacts := False;
  VH_RT.RTCreatInfos := False;
  bConsult1 := false; bConsult2 := false; bModif1 := false; bModif2 := false;

  if (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = true) then
    begin
    TobAcces:=TOB.create ('GRC Confidentialité',NIL,-1);
    QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'"  and RTC_PRODUITPGI="GRC"',True,-1,'PROSPECTCONF',true) ;
    if Not QQ.EOF then
    begin
     TobAcces.LoadDetailDB('PROSPECTCONF','','',QQ, False) ;
     for i:=0 to TobAcces.detail.count-1 do
     begin
     if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CON' then
        begin
        Consult1 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
        bConsult1:=true;
        end

     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CO2' then
        begin
        Consult2 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
        bConsult2:=true;
        end

    else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='MOD' then
        begin
        Modif1 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
        bModif1:=true;
        end

     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='MO2' then
        begin
        Modif2 := TobAcces.detail[i].GetValue('RTC_SQLCONF') ;
        bModif2:=true;
        end

     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CPR' then
        begin
        if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
           VH_RT.RTCreatPropositions := true;
        end
     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CAC' then
        begin
        if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
           VH_RT.RTCreatActions := true;
        end
     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CCO' then
        begin
        if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
           VH_RT.RTCreatContacts := true;
        end
     else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CIC' then
        begin
        if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
           VH_RT.RTCreatInfos := true;
        end
     ;
     end;

     if Consult2 = '' then  VH_RT.RTConfWhereConsult := Consult1
     //mng 25/06/03 else VH_RT.RTConfWhereConsult := copy(Consult1,5,(Length(Consult1)-4)) + Consult2 ;
     else VH_RT.RTConfWhereConsult := ' AND ('+copy(Consult1,5,(Length(Consult1)-4)) + Consult2 +')' ;

     if Modif2 = '' then  VH_RT.RTConfWhereModif := Modif1
     else VH_RT.RTConfWhereModif := ' AND ('+copy(Modif1,5,(Length(Modif1)-4)) + Modif2 +')';
     { mng 04/12/2007 : cas inexistence des enreg, consultation seule }
     if bConsult1 or bConsult2 or bModif1 or bModif2 then
       VH_RT.DroitModifTiers := True;
    end;
    Ferme(QQ);
    TobExist:=TobAcces.FindFirst(['RTC_TYPECONF'], ['MOD'], False);
    if Assigned(TobExist) then VH_RT.RTExisteConfident := true;
    TobAcces.free;
    end;
  { idem fournisseur}
  VH_RT.RFConfWhereConsult:='';
  VH_RT.RFConfWhereModif:='' ; Consult1:='';Consult2:='';Modif1:='';Modif2:='';
  VH_RT.RFDroitModifTiers := False;
  VH_RT.RFExisteConfident := False;
  VH_RT.RFCreatContacts := False;
  VH_RT.RFCreatInfos := False;
  VH_RT.RFCreatActions := False;
  bConsult1 := false; bConsult2 := false; bModif1 := false; bModif2 := false;
  if (GetParamsocSecur('SO_RFCONFIDENTIALITE',False) = False) then exit;
  TobAcces:=TOB.create ('GRC Confidentialité',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'"  and RTC_PRODUITPGI="GRF"',True,-1,'PROSPECTCONF',true) ;
  if Not QQ.EOF then
  begin
   TobAcces.LoadDetailDB('PROSPECTCONF','','',QQ, False) ;
   for i:=0 to TobAcces.detail.count-1 do
   begin
   if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CON' then
   begin
      Consult1 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
      bConsult1 := true;
   end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CO2' then
   begin
      Consult2 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
      bConsult2 := true;
   end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='MOD' then
   begin
      Modif1 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
      bModif1 := true;
   end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='MO2' then
   begin
      Modif2 := TobAcces.detail[i].GetValue('RTC_SQLCONF');
      bModif2 := true;
   end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CCO' then
      begin
      if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
         VH_RT.RFCreatContacts := true;
      end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CIC' then
      begin
      if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
         VH_RT.RFCreatInfos := true;
      end
   else if TobAcces.detail[i].GetValue('RTC_TYPECONF')='CAC' then
      begin
      if TobAcces.detail[i].GetValue('RTC_VAL1') = '=' then
         VH_RT.RFCreatActions := true;
      end
     ;
   end;

   if Consult2 = '' then  VH_RT.RFConfWhereConsult := Consult1
   //mng 25/06/03 else VH_RT.RTConfWhereConsult := copy(Consult1,5,(Length(Consult1)-4)) + Consult2 ;
   else VH_RT.RFConfWhereConsult := ' AND ('+copy(Consult1,5,(Length(Consult1)-4)) + Consult2 +')' ;

   if Modif2 = '' then  VH_RT.RFConfWhereModif := Modif1
   else VH_RT.RFConfWhereModif := ' AND ('+copy(Modif1,5,(Length(Modif1)-4)) + Modif2 +')';
   { mng 04/12/2007 : cas inexistence des enreg, consultation seule }
   if bConsult1 or bConsult2 or bModif1 or bModif2 then
     VH_RT.RFDroitModifTiers := True;
  end;
  Ferme(QQ);
  TobExist:=TobAcces.FindFirst(['RTC_TYPECONF'], ['MOD'], False);
  if Assigned(TobExist) then VH_RT.RFExisteConfident := true;
  TobAcces.free;
end;

Procedure InitLaVariableRT ;
  function GetSqlModActions: string;
  begin
    Result := 'SELECT * FROM ACTIONSGENERIQUES WHERE RAG_OPERATION="'+ModeleAction+'"';
  end;
begin
  {$IFNDEF EAGLSERVER}
     VH_RT:=LaVariableRT.Create ;
  {$ENDIF !EAGLSERVER}
  VH_RT.CTISeria := true; // En attendant la sérialisation du module
  VH_RT.RTModActionsCTI := tMemoryTob.Create('_MODACTIONS_', GetSqlModActions);
  VH_RT.AncienModeCTI:=false;
end ;

Procedure LibereLaVariableRT ;
Begin
  {$IFNDEF EAGLSERVER}
  if Assigned(VH_RT) then
  {$ENDIF !EAGLSERVER}
  begin
    VH_RT.TobChampsPro.Free;
    VH_RT.TobChampsProMul.Free;
    VH_RT.TobParamCl.Free;
    {$IFDEF GIGI}    //MCD 04/07/2005
      VH_RT.TobChampsDPMul.Free;
    {$ENDIF GIGI}
    VH_RT.TobTypesAction.Free;
    VH_RT.TobTypesChainage.Free;
    VH_RT.TobParamPlanning.Free;
    VH_RT.TobParamEtatPlan.Free;
    FreeAndNil(VH_RT.RTModActionsCTI);
    {$IFNDEF EAGLSERVER}
      VH_RT.Free ; VH_RT:=Nil ;
    {$ENDIF !EAGLSERVER}
  end
End;

Procedure ChargeParamsRT ;
BEGIN
OnChangeUserGRC ;
VH_RT.TobChampsPro := TMemoryTob.Create ('_LADESCRO_', '', ChargeTobRT);
ChargeTypesActions;
ChargeTypesChainages;
end;
{$ifdef GIGI} //MCD 04/07/2005
Procedure ChargeTobDP (Ismul: boolean=true);
var Q : TQuery;
  NumBase: integer;
begin
Numbase := 0;
if VH_RT.TobChampsDPMul<>Nil then VH_RT.TobChampsDPMul.Free ;
VH_RT.TobChampsDPMul:=TOB.Create('',Nil,-1) ;
Q:=OPENSQL('SELECT SO_VERSIONBASE FROM SOCIETE ',TRUE,-1, '', True) ;
If Not Q.Eof Then NumBase:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
if NumBase < 726 then exit; //mcd 18/01/2006 pour ne pas planter avec exe difssuion sur base 714
  //on doit tout prendre... peu de choses
	//si Ismul on est en saisie paramétrage, il faut prendre dans ordre clé
if Ismul then Q := OpenSQL('SELECT * FROM AFINFOGRCDP  WHERE ADP_TYPEINFO="DP" ORDER BY ADP_ONGLET ASC,ADP_PANEL DESC', true,-1,'AFINFOGRCDP',true)
 else Q := OpenSQL('SELECT * FROM AFINFOGRCDP  WHERE ADP_TYPEINFO="DP" ORDER BY ADP_ONGLET,ADP_PANEL', true,-1,'AFINFOGRCDP',true);
if Not Q.EOF then
  VH_RT.TobChampsDPMul.LoadDetailDB('AFINFOGRCDP','','',Q,False) ;
Ferme(Q);
end;
{$endif}

Procedure ChargeTobRT ;
var
  i : integer;
begin

if VH_RT.TobChampsProMul<>Nil then VH_RT.TobChampsProMul.Free ;
VH_RT.TobChampsProMul:=TOB.Create('',Nil,-1) ;

if VH_RT.TobParamCl<>Nil then VH_RT.TobParamCl.Free ;
VH_RT.TobParamCl:=TOB.create ('Param Cl',NIL,-1);
VH_RT.TobParamCl.LoadDetailFromSql ('Select * from COMMUN Where CO_TYPE="RPC" and CO_LIBRE="X"');

if VH_RT.TobParamCl.detail.count > 0 then
   for i := 0 to (VH_RT.TobParamCl.detail.count-1) do
       ChargeTobRTFille (VH_RT.TobParamCl.detail[i]);
{$ifdef GIGI}   //MCD 04/07/2005
 CHargeTobDP;
{$ENDIF GIGI}
end;

Procedure ChargeTypesActions ;
begin
// Types d'actions
VH_RT.TobTypesAction := TMemoryTob.Create ('RPA', GetSqlRPA);
end;

Procedure ChargeTypesChainages ;
begin
VH_RT.TobTypesChainage := TMemoryTob.Create ('RPG', GetSqlRPG);
end;


Procedure ChargeTobRTFille ( TobParam : Tob );
var j,k : integer;
    FieldNameFrom,FieldNameTo,codeFichier : String;
    TobChampsProFille,TobCli,TobPetiteFille : TOB ;
    Q : TQuery ;
begin
    codeFichier:=TobParam.GetString('CO_CODE');
    if codeFichier <> '0' then
      if not GetParamSocSecur('SO_RTGESTINFOS00'+codeFichier,true) then exit;
    TobCli:=TOB.Create('le client',Nil,-1) ;
    TobChampsProFille:=TOB.Create('la fille descro',VH_RT.TobChampsPro,-1) ;
    TobChampsProFille.Dupliquer(TobParam, False, True);
    if codeFichier <> '0' then
       TobChampsProFille.LoadDetailDB('RTINFOSDESC','"'+codeFichier+'"','RDE_DESC,RDE_ONGLET,RDE_PANEL',Nil,True)
    else
    // cas particulier du paramétrage client/prospect mémorisé dans CHAMPSPRO
        begin
        TobCli.LoadDetailDB('CHAMPSPRO','','RCL_ONGLET,RCL_PANEL',Nil,True) ;
        for j := 0 to TobCli.Detail.Count-1 do
            begin
            TobPetiteFille:=TOB.Create('',TobChampsProFille,j) ;
            TobPetiteFille.AddChampSupValeur ('RDE_DESC', '0');
            for k :=1 to TobCli.Detail[j].NbChamps do
                begin
                FieldNameFrom := TobCli.Detail[j].GetNomChamp(k);
                FieldNameTo := 'RDE_' + copy(FieldNameFrom,5,length(FieldNameFrom)) ;
                TobPetiteFille.AddChampSupValeur (FieldNameTo, TobCli.Detail[j].GetValue(FieldNameFrom));
                end;
            end;
        end;
    TobCli.Free;

    TobCli:=TOB.Create('le client',Nil,-1) ;

    TobChampsProFille:=TOB.Create('la fille descro',VH_RT.TobChampsProMul,-1) ;
    TobChampsProFille.Dupliquer(TobParam, False, True);
    if codeFichier <> '0' then
        begin
        Q := OpenSQL('SELECT * FROM RTINFOSDESC where RDE_DESC="'+codeFichier+'" ORDER BY RDE_ONGLET ASC,RDE_PANEL DESC', true,-1,'RTINFOSDESC',true);
        TobChampsProFille.LoadDetailDB('RTINFOSDESC','','',Q,False) ;
        Ferme(Q);
        end
    else
    // cas particulier du paramétrage client/prospect mémorisé dans CHAMPSPRO
        begin
        Q := OpenSQL('SELECT * FROM CHAMPSPRO ORDER BY RCL_ONGLET ASC,RCL_PANEL DESC', true,-1,'CHAMPSPRO',true);
        TobCli.LoadDetailDB('CHAMPSPRO','','',Q,False) ;
        Ferme(Q);
        for j := 0 to TobCli.Detail.Count-1 do
            begin
            TobPetiteFille:=TOB.Create('',TobChampsProFille,j) ;
            TobPetiteFille.AddChampSupValeur ('RDE_DESC', '0');
            for k :=1 to TobCli.Detail[j].NbChamps do
                begin
                FieldNameFrom := TobCli.Detail[j].GetNomChamp(k);
                FieldNameTo := 'RDE_' + copy(FieldNameFrom,5,length(FieldNameFrom)) ;
                TobPetiteFille.AddChampSupValeur (FieldNameTo, TobCli.Detail[j].GetValue(FieldNameFrom));
                end;
            end;
        end;
    TobCli.Free;
end;

Function RTGetResponsable : string;
var Q : TQuery;
begin
result:='';
Q:=OpenSQL('Select ARS_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_UTILASSOCIE="'+V_PGI.User+'"',true,-1,'RESSOURCE',true);
if Not Q.EOF then result:=Q.FindField('ARS_RESSOURCE').asstring+';'+Q.FindField('ARS_LIBELLE2').asstring+' '+Q.FindField('ARS_LIBELLE').asstring;
ferme(Q);
end;

Procedure RTGetCommercial ;
var Q : TQuery;
begin
VH_RT.RTCodeCommercial := ''; VH_RT.RTNomCommercial := '';
Q:=OpenSQL('SELECT GCL_COMMERCIAL,GCL_LIBELLE FROM COMMERCIAL WHERE GCL_UTILASSOCIE="'+V_PGI.User+'"',True,-1,'COMMERCIAL',true);
if Not Q.EOF then
begin
VH_RT.RTCodeCommercial := Q.fields[0].asString ;
VH_RT.RTNomCommercial  := Q.fields[1].asString ;
end;
ferme(Q);
end;

Procedure ChargeParamPlanning ;
begin
if VH_RT.TobParamPlanning<>Nil then VH_RT.TobParamPlanning.Free ;
VH_RT.TobParamPlanning := TMemoryTob.Create ('RPP', GetSqlRPP);
end;

Procedure ChargeParamEtatPlan ;
begin
if VH_RT.TobParamEtatPlan<>Nil then VH_RT.TobParamEtatPlan.Free ;
VH_RT.TobParamEtatPlan:=TMemoryTob.Create ('REP', GetSqlREP);
end;

{$IFDEF EAGLSERVER}
function VH_RT: LaVariableRT;
begin
  Result := LaVariableRT(RegisterVHSession('VH_RT', LaVariableRT))
end;
{$ENDIF EAGLSERVER}

end.

