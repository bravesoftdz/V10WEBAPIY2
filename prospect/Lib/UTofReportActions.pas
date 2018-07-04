unit UTofReportActions;

interface
uses  Controls,Classes,forms,sysutils,HMsgBox,Hstatus,M3FP,
      HCtrls,HEnt1,UTOF,ParamSoc,UtilRT,UTob,EntRT,HQry,UtilAction,Yplanning,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul;
{$ELSE}
      HDB,db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,Mul;
{$ENDIF}

Const
    Prevue : string = 'PRE';
    Annulee : string = 'ANU';

Type
     TOF_ReportActions = Class (TOF)
    private
           stWhere,CodeAuxiliaire,Libelle,CodeTiers,TypeAction : string ;
           NoAction : Integer;
           DateReport,DateRappel,DateEch : TDateTime;
           HeureAct : TDateTime;
           Dureeact : double;
        function ControleEstLibre : Boolean;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure RTReportActions;
        procedure SetAllReportAnnul;
        procedure SetReportAnnul;
     END ;

Function RTLanceFiche_ReportActions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
Function CalculDateRappel (GestRappel,DelaiRappel : string;HeureAction,DateAction,DateRappel : TDateTime) : TDateTime;
Function CalculDateEch (TypeAction,Chainage : string;Noligne:Integer;DateAction,DateEch : TDateTime) : TDateTime;

const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}        'Impossible de reporter ou d''annuler l''action %s du tiers %s'
          {2}        ,'Acceptez-vous la surcharge?#13#10   - Si vous répondez "OUI" toutes les actions sélectionnées sont reportées.#13#10   - Si vous répondez "NON" les actions dont les intervenants ne sont pas disponibles ne sont pas reportées'
          );

implementation

procedure TOF_ReportActions.OnArgument(Arguments : String ) ;
var DateHier : TDateTime;
    soRtactgestech: boolean;
begin
inherited ;

if (TFMul(Ecran).name = 'RTACTIONS_REPORT') then soRtactgestech := GetParamsocSecur('SO_RTACTGESTECH',False)
else soRtactgestech := GetParamsocSecur('SO_RFACTGESTECH',False);
if soRtactgestech = FALSE then
 begin
 SetControlEnabled('RAC_DATEECHEANCE',FALSE) ;
 SetControlEnabled('RAC_DATEECHEANCE_',FALSE) ;
 SetControlEnabled('TRAC_DATEECHEANCE',FALSE) ;
 SetControlEnabled('TRAC_DATEECHEANCE_',FALSE) ;
 end;
SetControlText('RAC_ETATACTION',Prevue);
SetControlText('RAC_INTERVENANT',VH_RT.RTResponsable);
SetControlText ('DATEJOUR',DateToStr(Date));
DateHier := PlusDate(Date,-1,'J');
SetControlText ('RAC_DATEACTION_',DateToStr(DateHier));
Arguments := uppercase(Arguments);
if Arguments = 'FICHEINFOS' then SetControlEnabled('RAC_INTERVENANT',FALSE);
{$IFDEF GRCLIGHT}
  if (TFMul(Ecran).name = 'RTACTIONS_REPORT') and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
{$ENDIF GRCLIGHT}
end;


procedure TOF_ReportActions.OnLoad;
begin
inherited;
if (TFMul(Ecran).name = 'RTACTIONS_REPORT') then setControlText('XX_WHERE',RTXXWhereConfident('CON'))
else setControlText('XX_WHERE',RTXXWhereConfident('CONF')) ;
end;

procedure tof_ReportActions.RTReportActions;
var  F : TFMul;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
     Q : THQuery;
     QQ : TQuery;
     i : integer;
     NotModif,ModifOk : Boolean;
     TobTypActEncours : Tob;
     Select,pgi : String;
     Rep : integer;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;

if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then exit ;
if (WriteYPLanning ('RAI')) and (GetControlText ('REPORT') = 'X') then Rep := PGIAsk(TexteMessage[2],'')
else Rep := mrYes;

{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end else
     F.Fliste.AllSelected := true;
{$ENDIF}
if copy(TFMul(Ecran).name,1,2) = 'RF' then pgi:='GRF' else pgi:='GRC';
DateReport := StrToDate(GetControlText('DATEREPORT'));
NotModif := False;

L:= F.FListe;
Q:= F.Q;

VH_RT.TobTypesAction.Load;

if L.AllSelected then
   begin
{   stWhere:=RecupWhereCritere(Pages);
   if (Transactions (SetAllReportAnnul,3) <> oeOK) then PGIBox('Impossible de reporter ou d''annuler toutes les actions ', F.Caption); }
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
      begin
      MoveCur(False);
      if ((TFMul(Ecran).name = 'RTACTIONS_REPORT') and RTDroitModifActions('',F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_INTERVENANT').asstring)) or
         ((TFMul(Ecran).name = 'RFACTIONS_REPORT') and RTDroitModifActionsF('',F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_INTERVENANT').asstring)) then
      begin
{debut}
        { accès paramètres du type d'action pour actions modifiable que par le responsable :
          ne doit pas être modifiable dans ce cas, sauf s'il s'agit du même responsable }
        ModifOK:=true;
        TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[F.Q.FindField('RAC_TYPEACTION').asstring,'---',0],TRUE) ;
        if TobTypActEncours <> Nil then
        begin
          if TobTypActEncours.GetValue('RPA_MODIFRESP')='X' then
          begin
            if not RTDroitModifTypeAction(pgi) then
            begin
              Select := 'SELECT ARS_UTILASSOCIE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ F.Q.FindField('RAC_INTERVENANT').asstring+'"';
              QQ:=OpenSQL(Select, True);
              if not QQ.Eof then
                 if QQ.FindField('ARS_UTILASSOCIE').asString <> V_PGI.User then
                     ModifOK:=False;
              Ferme(QQ) ;
            end;
          end;
        end
        else
           ModifOK:=False;

        if ModifOK then
        begin
{fin}
         CodeAuxiliaire:=F.Q.FindField('RAC_AUXILIAIRE').asstring ;
         CodeTiers:=F.Q.FindField('RAC_TIERS').asstring ;
         NoAction:=F.Q.FindField('RAC_NUMACTION').asinteger ;
         Libelle:=F.Q.FindField('RAC_LIBELLE').asstring ;
         TypeAction:=F.Q.FindField('RAC_TYPEACTION').asstring ;
         Heureact := 0;
         if F.Q.FindField('RAC_HEUREACTION').AsDateTime <> IDate1900 then Heureact := F.Q.FindField('RAC_HEUREACTION').AsDateTime;
         DureeAct:=F.Q.FindField('RAC_DUREEACTION').asfloat ;
         if (GetControlText ('REPORT') = 'X') and (Rep = mrNo) and (RTActEstPlanifiable (F.Q.FindField('RAC_TYPEACTION').asstring) = True) then ModifOK := ControleEstLibre;
         if ModifOK then
         begin
           if GetControlText ('REPORT') = 'X' then
              begin
              DateRappel := CalculDateRappel (F.Q.FindField('RAC_GESTRAPPEL').asstring,F.Q.FindField('RAC_DELAIRAPPEL').asstring,
                                              F.Q.FindField('RAC_HEUREACTION').AsDateTime,DateReport,
                                              F.Q.FindField('RAC_DATERAPPEL').AsDateTime);
              DateEch := CalculDateEch (F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_CHAINAGE').asstring,
                                              F.Q.FindField('RAC_NUMLIGNE').AsInteger,DateReport,
                                              F.Q.FindField('RAC_DATEECHEANCE').AsDateTime);
              end;
{           CodeAuxiliaire:=F.Q.FindField('RAC_AUXILIAIRE').asstring ;
           CodeTiers:=F.Q.FindField('RAC_TIERS').asstring ;
           NoAction:=F.Q.FindField('RAC_NUMACTION').asinteger ;
           Libelle:=F.Q.FindField('RAC_LIBELLE').asstring ;
           TypeAction:=F.Q.FindField('RAC_TYPEACTION').asstring ;
           Heureact := 0;
           if F.Q.FindField('RAC_HEUREACTION').AsDateTime <> IDate1900 then Heureact := F.Q.FindField('RAC_HEUREACTION').AsDateTime;
           DureeAct:=F.Q.FindField('RAC_DUREEACTION').asfloat ;  }
           if (Transactions (SetReportAnnul,3) <> oeOK) then PGIBox(format(TraduireMemoire(TexteMessage[1]),[Libelle,codetiers]), F.Caption);
         end
         else NotModif := True;
        end
        else NotModif := True;
      end
      else NotModif := True;
      Q.Next;
      end;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
      if ((TFMul(Ecran).name = 'RTACTIONS_REPORT') and RTDroitModifActions('',F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_INTERVENANT').asstring)) or
         ((TFMul(Ecran).name = 'RFACTIONS_REPORT') and RTDroitModifActionsF('',F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_INTERVENANT').asstring)) then
      begin
{debut}
        { accès paramètres du type d'action pour actions modifiable que par le responsable :
          ne doit pas être modifiable dans ce cas, sauf s'il s'agit du même responsable }
        ModifOK:=true;
        TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[F.Q.FindField('RAC_TYPEACTION').asstring,'---',0],TRUE) ;
        if TobTypActEncours <> Nil then
        begin
          if TobTypActEncours.GetValue('RPA_MODIFRESP')='X' then
          begin
            if not RTDroitModifTypeAction(pgi) then
            begin
              Select := 'SELECT ARS_UTILASSOCIE FROM RESSOURCE WHERE ARS_RESSOURCE = "'+ F.Q.FindField('RAC_INTERVENANT').asstring+'"';
              QQ:=OpenSQL(Select, True);
              if not QQ.Eof then
                 if QQ.FindField('ARS_UTILASSOCIE').asString <> V_PGI.User then
                     ModifOK:=False;
              Ferme(QQ) ;
            end;
          end;
        end
        else
           ModifOK:=False;

        if ModifOK then
        begin
{fin}
         CodeAuxiliaire:=F.Q.FindField('RAC_AUXILIAIRE').asstring ;
         CodeTiers:=F.Q.FindField('RAC_TIERS').asstring ;
         NoAction:=F.Q.FindField('RAC_NUMACTION').asinteger ;
         Libelle:=F.Q.FindField('RAC_LIBELLE').asstring ;
         TypeAction:=F.Q.FindField('RAC_TYPEACTION').asstring ;
         Heureact := 0;
         if F.Q.FindField('RAC_HEUREACTION').AsDateTime <> IDate1900 then Heureact := F.Q.FindField('RAC_HEUREACTION').AsDateTime;
         DureeAct:=F.Q.FindField('RAC_DUREEACTION').asfloat ;
         if (GetControlText ('REPORT') = 'X') and (Rep = mrNo) and (RTActEstPlanifiable (F.Q.FindField('RAC_TYPEACTION').asstring) = True) then ModifOK := ControleEstLibre;
         if ModifOK then
         begin
           if GetControlText ('REPORT') = 'X' then
              begin
              DateRappel := CalculDateRappel (F.Q.FindField('RAC_GESTRAPPEL').asstring,F.Q.FindField('RAC_DELAIRAPPEL').asstring,
                                              F.Q.FindField('RAC_HEUREACTION').AsDateTime,DateReport,
                                              F.Q.FindField('RAC_DATERAPPEL').AsDateTime);
              DateEch := CalculDateEch (F.Q.FindField('RAC_TYPEACTION').asstring,F.Q.FindField('RAC_CHAINAGE').asstring,
                                              F.Q.FindField('RAC_NUMLIGNE').AsInteger,DateReport,
                                              F.Q.FindField('RAC_DATEECHEANCE').AsDateTime);
              end;
{           CodeAuxiliaire:=F.Q.FindField('RAC_AUXILIAIRE').asstring ;
           CodeTiers:=F.Q.FindField('RAC_TIERS').asstring ;
           NoAction:=F.Q.FindField('RAC_NUMACTION').asinteger ;
           Libelle:=F.Q.FindField('RAC_LIBELLE').asstring ;
           TypeAction:=F.Q.FindField('RAC_TYPEACTION').asstring ;
           Heureact := 0;
           if F.Q.FindField('RAC_HEUREACTION').AsDateTime <> IDate1900 then Heureact := F.Q.FindField('RAC_HEUREACTION').AsDateTime;
           DureeAct:=F.Q.FindField('RAC_DUREEACTION').asfloat ; }
           if (Transactions (SetReportAnnul,3) <> oeOK) then PGIBox(format(TraduireMemoire(TexteMessage[1]),[Libelle,codetiers]), F.Caption);
         end
         else NotModif := True;
        end
        else NotModif := True;
      end
      else NotModif := True;
      end;
   L.ClearSelected;
   end;
if F.bSelectAll.Down then
    F.bSelectAll.Down := False;
FiniMove;
if NotModif =True then
  begin
  if Rep = mrNo then PGIBox('Certaines actions ne sont pas reportées ou annulées - Vous n''êtes pas autorisés à les modifier ou les intervenants ne sont pas disponibles', F.Caption)
  else PGIBox('Certaines actions ne sont pas reportées ou annulées - Vous n''êtes pas autorisés à les modifier', F.Caption);
  end;
end ;

// Fct plus utilisée
procedure tof_ReportActions.SetAllReportAnnul  ;
begin
if GetControlText ('REPORT') = 'X' then
     ExecuteSql('UPDATE ACTIONS SET RAC_DATEACTION="'+UsDateTime(DateReport)+'", '+
      'RAC_DATEMODIF="'+UsDateTime(Date)+'" '+StWhere)
else ExecuteSql('UPDATE ACTIONS SET RAC_ETATACTION="'+Annulee+'", '+
      'RAC_DATEMODIF="'+UsDateTime(Date)+'" '+StWhere);
end;

procedure tof_ReportActions.SetReportAnnul  ;
var TobRessources : Tob;
    Q    : TQuery;
    iDateDebut,iDateFin : TDateTime;
    Heures,Minutes : integer;
    i : integer;
begin
if GetControlText ('REPORT') = 'X' then
  begin
    ExecuteSql('UPDATE ACTIONS SET RAC_DATEACTION="'+UsDateTime(DateReport)+'", '+
       'RAC_DATERAPPEL="'+UsTime(DateRappel)+'", '+
       'RAC_DATEECHEANCE="'+UsTime(DateEch)+'", '+
       'RAC_DATEMODIF="'+UsDateTime(Date)+'", '+
       'RAC_UTILISATEUR="'+V_PGI.User+'" Where RAC_AUXILIAIRE="'+CodeAuxiliaire+
       '" and RAC_NUMACTION ='+ IntToStr(NoAction));
    if RTActEstPlanifiable (TypeAction) then
      begin
      TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
      Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+CodeAuxiliaire+'" AND RAI_NUMACTION ='+IntToStr(NoAction),True);
      try
       TobRessources.loaddetailDB('','','',Q ,false,False);
      finally
       ferme(Q);
      end;
      iDateDebut := DateReport + HeureAct;
      Heures:=trunc(Dureeact/60);
      Minutes:=trunc(Dureeact-(Heures*60));
      iDateFin := DateReport + HeureAct + EncodeTime(Heures,Minutes,0,0);
      RTMAJYPlanning (TobRessources,iDateDebut,iDateFin,'','');
      FreeAndNil (TobRessources);
      end;
  end
else
  begin
    ExecuteSql('UPDATE ACTIONS SET RAC_ETATACTION="'+Annulee+'", '+
       'RAC_DATEMODIF="'+UsDateTime(Date)+'", '+
       'RAC_UTILISATEUR="'+V_PGI.User+'" Where RAC_AUXILIAIRE="'+CodeAuxiliaire+
       '" and RAC_NUMACTION ='+ IntToStr(NoAction));
    if RTActEstPlanifiable (TypeAction) then
    begin
      TobRessources := TOB.Create ('Les LIENSACTIONS',NIL,-1);
      Q:=OpenSql('SELECT RAI_GUID FROM ACTIONINTERVENANT WHERE RAI_AUXILIAIRE = "'+CodeAuxiliaire+'" AND RAI_NUMACTION ='+IntToStr(NoAction),True);
      try
       TobRessources.loaddetailDB('','','',Q ,false);
      finally
       ferme(Q);
      end;
      for i:=0 to TobRessources.detail.count-1 do
        begin
        DeleteYPL ('RAI',TobRessources.Detail[i].GetString('RAI_GUID'));
        end;
      FreeAndNil (TobRessources);
    end;
  end;
end;

function Tof_ReportActions.ControleEstLibre : Boolean;
var TobControleYplanning,TF,TobRessources : Tob;
    Heures,Minutes : integer;
    i : integer;
    Q : TQuery;
    iDateDebut,iDateFin : TDateTime;
begin
Result := True;
TobControleYplanning := Tob.create('CONTROLEYPLANNING',Nil,-1) ;
TobRessources:=Tob.create('les LIENSACTIONS',Nil,-1) ;
Q:=OpenSql('SELECT RAI_RESSOURCE,RAI_GUID,ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAI_RESSOURCE WHERE RAI_AUXILIAIRE = "'+CodeAuxiliaire+'" AND RAI_NUMACTION ='+IntToStr(NoAction),True);
try
  TobRessources.loaddetailDB('','','',Q ,false,False);
finally
  ferme(Q);
end;
for i:=0 to TobRessources.detail.count-1 do
  begin
  TF:=Tob.create('#CONTROLE',TobControleYplanning,-1);
  TF.AddChampSupValeur('RESSOURCE',TobRessources.Detail[i].GetString('RAI_RESSOURCE') );
  TF.AddChampSupValeur('NOMRESSOURCE',TobRessources.Detail[i].GetString('ARS_LIBELLE') );
  TF.AddChampSupValeur('GUID',TobRessources.Detail[i].GetString('RAI_GUID'));
  TF.AddChampSupValeur('LIBRE','X');
  TF.AddChampSupValeur('MOTIF','');
  end;
if TobControleYplanning.Detail.Count <> 0 then
  begin
  iDateDebut := DateReport + HeureAct;
  Heures:=trunc(Dureeact/60);
  Minutes:=trunc(Dureeact-(Heures*60));
  iDateFin := DateReport + HeureAct + EncodeTime(Heures,Minutes,0,0);
  Result := ControleIsFreeYPL (TobControleYplanning,iDateDebut,iDateFin,False);
  end;
FreeAndNil (TobRessources);
FreeAndNil (TobControleYplanning);
end;

Function CalculDateRappel (GestRappel,DelaiRappel : string;HeureAction,DateAction,DateRappel : TDateTime) : TDateTime;
var HeureAct : TDateTime;
begin
Result := DateRappel;
//if  (GestRappel = 'X' ) and (HeureAction <> iDate1900 ) and (DelaiRappel <> '' ) then
if  (GestRappel = 'X' ) and (HeureAction <> iDate1900 ) then
    begin
    HeureAct:=DateAction+HeureAction;
    // calcul différent suivant que l'on traite des heures ou des jours.
    if DelaiRappel = '' then
       Result := HeureAct
    else
       if DelaiRappel < '024' then
          Result := HeureAct-EncodeTime(Valeuri(DelaiRappel), 0, 0, 0)
       else
          Result := PlusDate(HeureAct,(Valeuri(DelaiRappel) div 24) * (-1) ,'J');
    end;
end;

Function CalculDateEch (TypeAction,Chainage : string;Noligne:Integer;DateAction,DateEch : TDateTime) : TDateTime;
var TobTypActEncours : TOB;
begin
Result := DateEch;

   VH_RT.TobTypesAction.Load;

TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,Chainage,Noligne],TRUE) ;
if TobTypActEncours = Nil then TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,'---',0],TRUE) ;
if TobTypActEncours <> Nil then
  begin
     if TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0 then Result := RTCalculEch(DateAction,StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND'));
  end;
end;

procedure AGLRTReportActions(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is tof_ReportActions) then tof_ReportActions(TOTOF).RTReportActions else exit;
end;

Function RTLanceFiche_ReportActions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

Initialization
registerclasses([TOF_ReportActions]);
RegisterAglProc('RTReportActions',True,0,AGLRTReportActions);
end.
