{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 22/09/2015
Modifié le ... : 22/09/2015
Description .. : Gestion des événements Liés à une items du planning
Mots clefs ... : PLANNING
*****************************************************************}

unit UtilActionPlanning;

interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Graphics,
  HPlanning,
  StdCtrls,
  ExtCtrls,
  Controls,
  HCtrls,
  HStatus,
  MsgUtil,
  SysUtils,
  Hent1,
  dialogs,
  HMsgBox,
  HRichOLE,
  Paramsoc,
  HeureUtil,
  uJoursFeries,
	UAFO_Ferie,
  Windows,
  AglMail,
  MailOl,
  UDateUtils,
  PlanUtil,
  UTOB,
//
  Forms,
  ComCtrls,
{$IFDEF EAGLCLIENT}
  MaineAGL, UtilEagl,
  eMul,
{$ELSE}
  Fe_main,
  EdtREtat,
  mul,
{$ENDIF}
	AglInit,
  UIUtil,
  HPanel,
  Mask,
  HTB97,
  ParamDat,
  Lookup,
  BTsaisieDate,
  BTPUtil,
  Grids,
  Utof_VideInside,
  Menus, DBCtrls, HDB, HRichEdt, HCapCtrl,
  TntComCtrls, TntStdCtrls, TntGrids, TntExtCtrls;
//


type

  TTypeDateTrait = (TtDDebut,TtDFin);

  function  ControlCalendrier(DateDeb : TDateTime; HeureDeb, HeureFin : Double; Calendrier, Ressource : String) : Boolean;
  function  ConstitueDatePlanning (TobModelePlanning : TOB; UneDate : TDateTime; TypeDate : TTypeDateTrait) : TDateTime;
  Function  ControlEventAbs(DateDeb, DateFin : TDateTime; Ressource : String) : Boolean;
  Function  ControlEventCha(DateDeb, DateFin : TDateTime; Ressource : String) : Boolean;
  Function  ControlEventMat(DateDeb, DateFin : TDateTime; Materiel : String)  : Boolean;
  Function  ControlEventPla(TobItem : TOB) : Boolean;
  function  ControlParametres(HeureDeb, HeureFin : Double) : Boolean;
  function  ControleRessourceDispo (TobItem : Tob; CtrlCal : Boolean = True) : boolean;
  Procedure CreationActionGRC(Item : TOB);
  function  DeleteEvenementMateriel (Item : Tob; P : THplanningBTP) : boolean;
  Procedure FormatageDate (Var StDate   : TDateTime);
  Procedure FormatageHeure(Var StHeure  : Double);

  function GetPrefixeTob(TobItem : TOB) : string;

  function  IsModifiable(Item: TOB): boolean;

  procedure ModificationsActionGRC(Item : TOB);

implementation

uses  uBtpEtatPlanning ,
      DateUtils,
      TntClasses,
      Classes,
      UtilsParc,
      DB;


//Suppression d'un événement Planning Matériel
function  DeleteEvenementMateriel (Item : Tob; P : THplanningBTP) : boolean;
Var NumEvent    : Integer;
    Materiel    : string;
    Ressource   : string;
    Marequete   : string;
    TobDelete   : TOB;
begin

  Result := False;

  if Assigned(Item) then
  Begin
    NumEvent    := item.GetValue('BPL_IDEVENT');
    Materiel    := item.GetValue('BPL_MATERIEL');
    Ressource   := item.GetValue('BPL_RESSOURCE');
  end;

  If pgiAsk('Confirmez-vous la suppression de l''évènement ?')=MrNo then Exit;

  TobDelete := Item.FindFirst(['BPL_CODEEVENT'], [NumEvent], True);

  if (Item.GetValue('BPL_ORIGINEITEM') = 'PARCMAT') then Marequete :='DELETE FROM BTEVENTMAT WHERE BEM_IDEVENTMAT='+IntToStr(NumEvent)+' AND BEM_CODEMATERIEL="'+ Materiel +'"';

  if ExecuteSql (Marequete) = 0 then
  begin
    P.DeleteItem(Item);
    Result := True;
  end;


end;

function IsModifiable(Item: TOB): boolean;
var QQ          : TQuery;
    NumEventMat : Integer;
    CodeMateriel: string;
    codeAffaire : string;
begin

  result := False;

  if Item = nil then exit;

  if not assigned(Item) then exit;

  if Item.getValue('MODIFIABLE') <> 'X' then exit;

  CodeMateriel  := Item.GetString('BPL_MATERIEL');
  CodeAffaire   := Item.getValue('BPL_AFFAIRE');

  if (Item.GetValue('BPL_ORIGINEITEM') = 'ACTIONGRC') then
  begin
    QQ := OpenSql ('SELECT RAC_ETATACTION FROM ACTIONS WHERE RAC_AUXILIAIRE="' + Item.getValue('CLIENTGRC')+'" AND RAC_NUMACTION='+IntToStr(Item.getValue('ACTIONSGRC')),true,1,'',true);
    if not QQ.eof then
    begin
      if QQ.FindField('RAC_ETATACTION').AsString <> 'PRE' then result := false;
    end;
    ferme (QQ);
  end
  else if (Item.GetValue('BPL_ORIGINEITEM') = 'INTERV') then
  begin
    if CodeAffaire <> '' then
    BEGIN
      QQ := OpenSql ('SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire + '"',true,1,'',true);
      if not QQ.eof then
      begin
        if QQ.FindField('AFF_ETATAFFAIRE').AsString <> 'AFF' then
        begin
          Exit;
          Ferme(QQ);
        end;
      end;

      ferme (QQ);
    END;
  end;

	result := True;

end;

function ConstitueDatePlanning (TobModelePlanning : Tob; UneDate : TDateTime; TypeDate : TTypeDateTrait) : TDateTime;
var DateRec   : string;
    HeureRec  : Double;
    HeureCalc : string;
begin

  DateRec  := DateToStr(UneDate);
  HeureRec := TimeToFloat(UneDate);

  if Heurerec = 0 then
  begin
  	if TypeDate = TtdDebut then
    begin
  		HeureRec := TimeToFloat(TobModelePlanning.getvalue('HPP_HEUREDEBUT'), True);
    end else
    begin
  		HeureRec := TimeToFloat(TobModelePlanning.getvalue('HPP_HEUREFIN'), True);
    end;
  end;

  if (Heurerec = 0) and (TypeDate = TTdDebut) then HeureRec := GetParamSocSecur('SO_HEUREDEB', TimeToFloat(StrToTime('08:00:00')));
  if (Heurerec = 0) and (TypeDate = TTdFin)   then HeureRec := GetParamSocSecur('SO_HEUREFIN', TimeToFloat(StrToTime('18:00:00')));

  HeureCalc := FormatDateTime('hh:mm:ss', Heurerec);
  result    := Trunc(StrToDate(DateRec)) + StrToTime(HeureCalc);

end;


Procedure FormatageDate(Var StDate : TDateTime);
Var Y : Word;
    M : Word;
    J : Word;
    H : Word;
    MM: Word;
    S : Word;
    MS: Word;
begin

  DecodeDateTime(StDate, Y, M, J, H, MM, S, MS);

  StDate := EncodeDateTime(Y, M, J, H, MM, 00, 00);

end;

Procedure FormatageHeure(Var StHeure : Double);
Var Y : Word;
    M : Word;
    J : Word;
    H : Word;
    MM: Word;
    S : Word;
    MS: Word;
begin

  DecodeTime(StHeure, H, MM, S, MS);

  StHeure := EncodeTime(H, MM, 00, 00);

end;

function ControleRessourceDispo (TobItem : TOB; CtrlCal : Boolean = True) : boolean;
var TOBCalendar : TOB;
    //
    Prefixe   : string;
    //
    Materiel  : String;
    Ressource : String;
    Salarie   : String;
    TypeEvent : string;
    Calendrier: string;
    //
    DateDebEvt: TdateTime;
    DateFinEvt: TdateTime;
    //
    HeureDeb  : TTime;
    HeureFin  : TTime;
    //
    Y, M, J     : Word;
begin

  // par defaut elle est dipo
	result := False;
  //
  Prefixe := GetPrefixeTob(TobItem);

  If Prefixe = '' then Exit;
  //
  //On charge des zones intermediaires de la TobItem....
  Materiel  := TobItem.GetValue(Prefixe + '_MATERIEL');
  Ressource := TobItem.GetValue(Prefixe + '_RESSOURCE');
  Salarie   := TobItem.GetValue(Prefixe + '_SALARIE');
  Calendrier:= TobItem.GetValue(Prefixe + '_STANDCALEN');

  DateDebEvt:= TobItem.GetValue(Prefixe + '_DATEDEB');
  DecodeDate(DateDebEvt, Y, M, J);
  DateDebEvt := EncodeDate(Y,M,J);

  DateFinEvt:= TobItem.GetValue(Prefixe + '_DATEFIN');
  DecodeDate(DatefinEvt, Y, M, J);
  DateFinevt := EncodeDate(Y,M,J);

  HeureDeb  := TobItem.GetValue(Prefixe + '_HEUREDEB');
  HeureFin  := TobItem.GetValue(Prefixe + '_HEUREFIN');

  TypeEvent := TobItem.GetValue(Prefixe + '_ORIGINEITEM');

  //On vérifie dans un premier temps si la ressource ne dispose pas d'un événement sur la période...
  if (TypeEvent = 'PARCMAT') And (not ControlEventMat(DateDebEvt, DateFinEvt, Materiel)) then
  Begin
    PgiError ('Impossible : La ressource sélectionnée se trouve déjà sur un évènement Parc/matériel');
    Exit;
  end;

  if (TypeEvent = 'INTERV') and (not ControlEventPla(TobItem)) then
  Begin
    PgiError ('Impossible : La ressource sélectionnée se trouve déjà sur un évènement Intervention');
    Exit;
  end;

  if (TypeEvent = 'CHANTIER') And (not ControlEventCha(DateDebEvt, DateFinEvt, Ressource)) then
  Begin
    PgiError ('Impossible : La ressource sélectionnée se trouve déjà sur un évènement Chantier');
    Exit;
  end;

  //On devrait vérifier si la ressource n'est pas en congé ou en maladie (Abs)
  if (TypeEvent = 'INTERV') and (not ControlEventAbs(DateDebEvt, DateFinEvt, Salarie)) then
  Begin
    PgiError ('Impossible : La ressource sélectionnée est en absence de Paie');
    Exit;
  end;

  //On vérifie si la ressource a un calendrier si les date et les heures sont cohérentes avec ce calendrier
  if CtrlCal then
  begin
    if Calendrier <> '' then
    begin
      if not ControlCalendrier(DateDebEvt, HeureDeb, HeureFin, Calendrier, Ressource) then
      begin
        PgiError ('Impossible : La plage horaire ne correspond pas au calendrier de la ressource sélectionnée');
        Exit;
      end;
    end
    else
    begin
      if not ControlParametres(HeureDeb, HeureFin) then
      Begin
        PgiError ('Impossible : La plage horaire ne correspond pas au paramètres société');
        Exit;
      end;
    end;
  end;

  //On vérifie si date de début et date de fin ne sont pas des jours fériés
  if AglJoursFeries(DateDebEvt) then
  Begin
    if PgiAsk('Confirmez-vous le début sur un jour férié ?', 'Erreur de Date') = MrNo then exit;
  end;

  if AglJoursFeries(DateFinEvt) then
  Begin
    if PgiAsk('Confirmez-vous la fin sur un jour férié ?', 'Erreur de Date') = MrNo then  exit;
  end;

    //On vérifie si date de début et date de fin ne sont pas des Week-End
  if (DayOfWeek(DateDebEvt) = 1) Or (DayOfWeek(DateDebEvt) = 7) then
  begin
    if PgiAsk('Confirmez-vous le début sur un week-end ?', 'Erreur de Date') = MrNo then exit;
  end;

  if (DayOfWeek(DateFinEvt) = 1) Or (DayOfWeek(DateFinEvt) = 7) then
  begin
    if PgiAsk('Confirmez-vous la fin sur un week-end ?', 'Erreur de Date') = MrNo then  exit;
  end;

  //Vérification si date de début correspond à un jour férié ou non travaillé
  if IsDateClosed (DateDebEvt) then
  begin
    if PgiAsk('Confirmez-vous le début sur un jour non travaillé', 'Erreur de Date') = MrNo then  exit;
  end;
  //
  //Vérification si date de fin correspond à un jour férié ou non travaillé
  if IsDateClosed (DateFinEvt) then
  begin
    if PgiAsk('Confirmez-vous la fin sur un jour non travaillé', 'Erreur de Date') = MrNo then  exit;
  end;

  //On Vérifie que la date de début ne soit pas > à la date de fin
  if DateDebEvt > DateFinEvt then
  Begin
    PGIBox('La date de début est supérieure à la date de fin !', 'Erreur de Date');
    Exit;
  end;

  //On Vérifie que la date de début ne soit pas = à la date de fin
  {if DateDeb = DateFin then
  Begin
    PGIBox('La date de début et la date de fin sont égales !', 'Erreur de Date');
    exit;
  end;}

	result := True;

end;

function GetPrefixeTob(TobItem : TOB) : string;
begin

  result := '';

  if TobItem.nomtable = 'BTEVENPLAN'      then result := 'BEP' else
  if TobItem.nomtable = 'BTEVENEMENTPLA'  then result := 'BPL' else
  if TobItem.nomtable = 'BTEVENTCHA'      then result := 'BEC' else
  if TobItem.nomtable = 'BTEVENTMAT'      then result := 'BEM' else
  if TobItem.nomtable = ''                then result := ''    else result := '';

end;


Function ControlEventAbs(DateDeb, DateFin : TDateTime; Ressource : String) : Boolean;
Var StSQL         : string;
    QQ            : TQuery;
begin

  Result := True;

  if Ressource = '' then exit;

  StSQL := 'SELECT * FROM ABSENCESALARIE  WHERE PCN_SALARIE="' + Ressource + '" ';
  StSQL := StSQL + '  AND (PCN_DATEDEBUT  BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '" ';
  StSQL := StSQL + '   OR  PCN_DATEFIN    BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '") ';

  QQ := OpensQL(StSQL, False,-1,'',True);

  Result := QQ.eof;

  Ferme(QQ);

end;

Function ControlEventMat(DateDeb, DateFin : TDateTime; Materiel : String) : Boolean;
Var StSQL         : string;
    QQ            : TQuery;
begin

  Result := True;

  if Materiel = '' then exit;

  StSQL := 'SELECT * FROM BTEVENTMAT WHERE BEM_CODEMATERIEL="' + Materiel + '" ';
  StSQL := StSQL + '  AND (BEM_DATEDEB  BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '" ';
  StSQL := StSQL + '   OR  BEM_DATEFIN  BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '") ';

  QQ := OpensQL(StSQL, False,-1,'',True);

  Result := QQ.eof;

  Ferme(QQ);

end;

Function ControlEventPla(TobItem : TOB) : Boolean;
Var StSQL         : string;
    QQ            : TQuery;
    DateDeb       : TDateTime;
    Datefin       : TDatetime;
    HeureDeb      : TTime;
    HeureFin      : TTime;
begin

  Result := True;

  if TobItem.GetValue('BPL_RESSOURCE') = '' then exit;

  DateDeb  := TobItem.getValue('BPL_DATEDEB');
  DateFin  := TobItem.getValue('BPL_DATEFIN');
  HeureDeb := TobItem.getValue('BPL_HEUREDEB');
  HeureFin := TobItem.getValue('BPL_HEUREFIN');

  DateDeb  := DateDeb + HeureDeb;
  DateFin  := DateFin + HeureFin;


  StSQL := 'SELECT * FROM BTEVENPLAN WHERE BEP_RESSOURCE="' + TobItem.GetValue('BPL_RESSOURCE')  + '" ';

  If TobItem.GetValue('BPL_CODEEVENT') <> '' then
    StSQl := StSQL + '  AND BEP_CODEEVENT <> "' + TobItem.GetValue('BPL_IDEVENT') + '"';

  StSQL := StSQL + '  AND (BEP_DATEDEB  BETWEEN "' + USDATETIME(DateDeb) + '" ';
  StSQL := StSQL + '  AND "' + USDATETIME(DateFin) + '" ';
  StSQL := StSQL + '   OR  BEP_DATEFIN  BETWEEN "' + USDATETIME(DateDeb) + '" ';
  StSQL := StSQL + '  AND "' + USDATETIME(DateFin) + '") ';

  QQ := OpensQL(StSQL, False,-1,'',True);

  Result := QQ.eof;

  Ferme(QQ);

end;

Function ControlEventCha(DateDeb, DateFin : TDateTime; Ressource : String) : Boolean;
Var StSQL         : string;
    QQ            : TQuery;
    CodeRessource : String;
begin

  Result := True;

  if Ressource = '' then exit;

  StSQL := 'SELECT * FROM BTEVENTCHA WHERE BEC_RESSOURCE="' + Ressource + '" ';
  StSQL := StSQL + '  AND (BEC_DATEDEB  BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '" ';
  StSQL := StSQL + '   OR  BEC_DATEFIN  BETWEEN "' + USDATETIME(DateDeb) + '" AND "' + USDATETIME(DateFin) + '") ';

  QQ := OpensQL(StSQL, False,-1,'',True);

  Result := QQ.eof;

  Ferme(QQ);

end;

function ControlCalendrier(DateDeb : TDateTime; HeureDeb, HeureFin : Double; Calendrier, Ressource : String) : Boolean;
Var NoJour      : String;
    StSQL       : String;
    QQ          : TQuery;
    TobCalendar : TOB;
    StMonHeure  : String;
    AMDebut     : TTime;
    AMFin       : TTime;
    PMDebut     : TTime;
    PMFin       : TTime;
    HdebParam   : TTime;
    HFinParam   : TTime;
    H, M, S, Ms : Word;
begin

  Result := False;

  //chargement du N° du jour pour recherche dans calendrier
  NoJour      := IntToStr(DayOfWeek(DateDeb)-1);

  //Vérification si Calendrier Particulier à la Ressource
  StSQL := 'SELECT * FROM CALENDRIER WHERE ACA_JOUR  =' + NoJour;
  StSQL := StSQL + ' AND ACA_DATE <="' + UsDateTime(DateDeb) + '"';
  StSQL := StSQL + ' AND ACA_STANDCALEN="' + Calendrier + '"';
  StSQL := StSQL + ' AND ACA_RESSOURCE="' + Ressource + '"';

  QQ := OpenSQL(StSQL, true,-1,'',true);

  if QQ.eof then
  begin
    Ferme(QQ);
    StSQL := 'SELECT * FROM CALENDRIER WHERE ACA_JOUR  =' + NoJour;
    StSQL := StSQL + ' AND ACA_DATE <="' + UsDateTime(DateDeb) + '"';
    StSQL := StSQL + ' AND ACA_STANDCALEN="' + Calendrier + '"';
    StSQL := StSQL + ' AND ACA_RESSOURCE="***"';
    QQ := OpenSQL(StSQL, true,-1,'',true);
    if QQ.eof then
    begin
      Result := True;
      Ferme(QQ);
      Exit;
    end;
  end;

  TOBcalendar := TOB.Create('CALENDRIER', nil, -1);
  TOBcalendar.SelectDB('CALENDRIER', QQ);

  Ferme(QQ);

  //On vérifie si pas un jour férié ou non travaillé...
  if TobCalendar.GetBoolean('ACA_FERIETRAVAIL') then
  begin
    FreeAndNil(TobCalendar);
    Exit;
  end;

  //Chargement des zone d'heure....
  StMonHeure:= TimeToStr(TobCalendar.GetValue('ACA_HEUREDEB1')/24);
  AMDebut   := StrToTime(StMonHeure);
  DecodeTime(AMDebut, H,M,S,Ms);
  AMDebut   := EncodeTime(H,M,0,0);

  StMonHeure:= TimeToStr(TobCalendar.GetValue('ACA_HEUREFIN1')/24);
  AMfin     := StrToTime(StMonHeure);
  DecodeTime(AMfin, H,M,S,Ms);
  AMfin     := EncodeTime(H,M,0,0);

  StMonHeure:= TimeToStr(TobCalendar.GetValue('ACA_HEUREDEB2')/24);
  PMDebut   := StrToTime(StMonHeure);
  DecodeTime(PMDebut, H,M,S,Ms);
  PMDebut   := EncodeTime(H,M,0,0);

  StMonHeure:= TimeToStr(TobCalendar.GetValue('ACA_HEUREFIN2')/24);
  PMFin     := StrToTime(StMonHeure);
  DecodeTime(PMFin, H,M,S,Ms);
  PMFin     := EncodeTime(H,M,0,0);

  HdebParam := AmDebut;
  HFinParam := PMFin;

  if HeureDeb < AmDebut then
  begin
    //HeureDeb := AMDebut;
    FreeAndNil(TobCalendar);
    Exit;
  end;

  if (HeureDeb > AMFin) and (HeureDeb < PMDebut) then
  begin
    FreeAndNil(TobCalendar);
    Exit;
    //HeureDeb := AMFin;
  end;

  if HeureDeb > PMFin   then
  Begin
    //HeureDeb := PMFin;
    FreeAndNil(TobCalendar);
    Exit;
  end;

  Result := True;

  FreeAndNil(TobCalendar);

end;

function ControlParametres(HeureDeb, HeureFin : Double) : Boolean;
Var AMDebut   : Double;
    AMFin     : Double;
    PMDebut   : Double;
    PMFin     : Double;
    HdebParam : Double;
    HFinParam : Double;
begin

  Result    := False;
  //
  AMDebut   := GetParamSocSecur('SO_BTAMDEBUT', GetParamSocSecur('SO_HEUREDEB', TimeToFloat(StrToTime('08:45:00'))));
  AMfin     := GetParamSocSecur('SO_BTAMFIN',   TimeToFloat(StrToTime('12:20:00')));

  PMDebut   := GetParamSocSecur('SO_BTPMDEBUT', GetParamSocSecur('SO_HEUREFIN', TimeToFloat(StrToTime('14:00:00'))));
  PMFin     := GetParamSocSecur('SO_BTPMFIN',   TimeToFloat(StrToTime('17:45:00')));

  HdebParam := GetParamSocSecur('SO_HEUREDEB',  TimeToFloat(StrToTime('08:45:00')));
  HFinParam := GetParamSocSecur('SO_HEUREFIN',  TimeToFloat(StrToTime('17:45:00')));

  FormatageHeure(HeureDeb);
  FormatageHeure(HeureFin);

  FormatageHeure(AMDebut);
  FormatageHeure(AMFin);
  FormatageHeure(PMDebut);
  FormatageHeure(PMFin);

  FormatageHeure(HdebParam);
  FormatageHeure(HFinParam);

  //On vérifie que l'heure de début et l'heure de fin sont comprise dans l'heure de début et l'heure de fin des plages 1 et 2 du calendrier
  if (HeureDeb < AMDebut) or (HeureDeb > PMFin) then Exit;

  if ((HeureDeb > AMFin) And (HeureDeb < PMDebut))   or
     ((HeureFin > AMFin) And (HeureFin < PMDebut)) then Exit;

  if (HeureFin < AMDebut) or (HeureFin > PMFin) then  Exit;

  Result    := True;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 16/03/2016
Modifié le ... :   /  /
Description .. : Gestion des action possible sur évènement action-GRC
Mots clefs ... :
*****************************************************************}
Procedure CreationActionGRC(Item : TOB);
Var NumRes  : string;
    Retour  : String;
begin

  Numres  := Item.GetString('ARS_RESSOURCE');

  Retour  := AGLLanceFiche('RT','RTACTIONS','','','MONOFICHE;ACTION=CREATION;RAC_INTERVENANT='+NumRes) ;

end;

procedure ModificationsActionGRC(Item : TOB);
var resultat    : string;
    Auxiliaire  : string;
    NumAct      : Integer;
    DateDeb     : TdateTime;
    Duree       : integer;
begin

  Auxiliaire  := Item.GetValue('BPL_TIERS');
  NumAct      := Item.GetValue('BPL_IDEVENT');

  DateDeb     := Item.GetValue('BPL_DATEDEB');
  Duree       := Item.GetValue('BPL_DUREE');

  resultat :=   AGLLanceFiche('RT','RTACTIONS','',Auxiliaire + ';' + IntToStr(NumAct),'ACTION=MODIFICATION;MODIFPLANNING') ;

  if resultat <> '' then
  begin
    if (pos('DELETE',resultat) > 0) then
    begin
      //if GetParamSocSecur ('SO_BTAVERTIRENMODIF',true) then EnvoieEltEmailFromGRC (Auxiliaire,NumAct,'D',Idate1900,0);
      //DeleteItemGRC (Item);
    end
    else
    begin
      //if GetParamSocSecur ('SO_BTAVERTIRENMODIF',true) then EnvoieEltEmailFromGRC (Auxiliaire,NumAct,'M',DateDeb,Duree,true);
      //ReChargeActionGRC (Auxiliaire,NumAct);
    end;
  end;

end;

end.
