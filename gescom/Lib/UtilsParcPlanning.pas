unit UtilsParcPlanning;

interface

uses  M3FP,
      StdCtrls,
      Controls,
      Classes,
      Forms,
      SysUtils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOB,
      UTOF,
      AglInit,
      Agent,
      EntGC,
      StrUtils,
      DateUtils,
      HeureUtil,
      ParamSoc,
      Constantes,
      UProcGen,
      UtilSaisieconso,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      HTB97;

    function CalculeDateEtireReduit(HeureFin : TTime; Cadencement : String): TTime;

    function DeplacementParcMateriel(Item, TobJFerie : tob; Materiel  : String) : Boolean;
    function DeplacementPlanCharge  (Item : TOB; Ressource : String; Cadencement : String) : boolean;

    Function DuplicationParcMateriel(Item, TobJFerie : TOB; Materiel : String) : boolean;
    Function DuplicationPlancharge  (Item : TOB; Ressource: String; Cadencement : String)       : boolean;
    function GestionDateItem(Item : TOB) : Boolean;
    Function MAJPlanCharge(Item : Tob)  : Boolean;
    function ReductionEtirementParcMateriel (Item : TOB; Cadencement : String)                  : boolean;
    function ReductionEtirementPlanCharge   (Item : TOB; Cadencement : String) : boolean;

    function MiseAJourParcMateriel(Item : TOB) : Boolean;

implementation

uses  PlanUtil,
      UtilsParc,
      BTPUtil,
      UdateUtils,
      UtilActionPlanning;


function GestionDateItem(Item : TOB) : Boolean;
Var HeureDeb        : TTime;
    HeureFin        : TTime;
    Duree           : Double;
    DureeMin        : Integer;
    Datedebut       : TDatetime;
    Datefinal       : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
    NbHJour         : Double;
begin

  Result := False;

  NbHJour     := ChargeParametreTemps;

  //on recharge l'heure de début, l'heure de fin et on recalcul la durée
  DateDebut := StrToDateTime(Item.GetValue('BPL_DATEAFFDEB'));
  DecodeDateTime(DateDebut, an, Mois, Jour, H, M, S, Msec);
  HeureDeb  := EncodeTime(H, M, 0, 0);

  Datefinal := StrtoDateTime(Item.GetValue('BPL_DATEAFFFIN'));
  DecodeDateTime(DateFinal, An, Mois, Jour, H, M, S, Msec);
  HeureFin  := EncodeTime(H, M, 0, 0);
  HeureFin  := CalculeDateEtireReduit (HeureFin, Cadencement);

  DateFinal := Trunc(DateFinal) + HeureFin;

  //Vérification cohérence...
  if not ControleCoherence (DateDebut,DateFinal, Duree, Item.GetString('BPL_MATERIEL'), Item.GetValue('BPL_IDEVENT')) then Exit;

  Item.PutValue('BPL_DATEDEB', DateTimeToStr(DateDebut));
  Item.PutValue('BPL_DATEFIN', DateTimeToStr(DateFinal));

  Item.PutValue('BPL_HEUREDEB', TimeToStr(HeureDeb));
  Item.PutValue('BPL_HEUREFIN', TimeToStr(HeureFin));
  //
  Duree := CalculDureeEvenement(DateDebut, DateFinal);
  if Duree <> 0 then Duree := Duree/60;
  Item.PutValue('BPL_DUREE', Duree);

  Result := True;

end;

function ReductionEtirementParcMateriel (Item : TOB; Cadencement : String) : boolean;
begin

  if not GestionDateItem(Item) then Exit;

  MiseAJourParcMateriel(Item); //UpdateParcMateriel (Materiel, NumEvent, DateD, DateF);

end;

function DeplacementParcMateriel (Item, TobJFerie : TOB; Materiel : String) : boolean;
var CodeMat         : String;
    Rep			 	      : String;
begin

  Result := True;
  //
  CodeMat		  := Item.GetValue('BPL_MATERIEL');
  //
  if not GestionDateItem(Item) then Exit;

  //Controle s'il n'y a pas de changement d'affectation
  if CodeMat <> Materiel then
  Begin
    if PgiAsk('Confirmez-vous le changement d''affectation ?', rep)=mrNo then
    Begin
      Item.PutValue('BPL_MATERIEL', Materiel);
    end;
  end;

  MiseAJourParcMateriel(Item);

end;

Function DuplicationParcMateriel(Item, TobJFerie : TOB; Materiel : String) : boolean;
Var TobMateriel : TOB;
    TobEventMat : TOB;
    STSQL       : string;
    QQ          : TQuery;
    FF          : String;
    PA          : Double;
    PR          : Double;
    PV          : Double;
    x           : Integer;
    NewNumEvent : Integer;
    MaterielFin : String;
    NumEvent    : Integer;
    Duree           : Double;
    DureeMin        : Integer;
    Datedebut       : TDatetime;
    Datefinal       : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
    CodeMat         : String;
    NbHJour         : Double;
begin

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for x := 1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  Result := false;

	if not IsModifiable(Item) then
  Begin
    PGIError('Cet événement n''est  pas modifiable', 'Duplication événement Parc/Matériel');
    exit;
  end;

  if not GetNumCompteur('BEM',iDate1900, NewNumEvent) then
  begin
    PGIError(TexteMessage[7], 'Duplication événement Parc/Matériel');
    exit;
  end;

  MaterielFin := Item.GetValue('BPL_MATERIEL');
  NumEvent    := Item.GetValue('BPL_IDEVENT');
  //
  if not GestionDateItem(Item) then Exit;
  //
  StSQL := 'SELECT * FROM BTEVENTMAT WHERE BEM_IDEVENTMAT=' + IntToStr(NumEvent);
  QQ    := OpenSQL(StSql,False,1,'',True);
  if QQ.Eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  TobEventMat := TOB.Create('BTEVENTMAT',nil, -1);
  TobEventMat.SelectDB('BTEVENTMAT',QQ);
  //
  TobEventMat.PutValue('BEM_IDEVENTMAT',    NewNumEvent);
  TobEventMat.PutValue('BEM_CODEMATERIEL',  MaterielFin);
  //
  TobEventMat.PutValue('BEM_DATEDEB',  Item.GetValue('BPL_DATEDEB')) ;
  TobEventMat.PutValue('BEM_DATEFIN',  Item.GetValue('BPL_DATEFIN')) ;
  TobEventMat.PutValue('BEM_HEUREDEB', Item.GetValue('BPL_HEUREDEB'));
  TobEventMat.PutValue('BEM_HEUREFIN', Item.GetValue('BPL_HEUREFIN'));
  TobEventMat.PutValue('BEM_NBHEURE',  Item.GetValue('BPL_DUREE'));

  TobMateriel := Tob.create('BTMATERIEL', nil, -1);
  if Materiel <> '' then ChargeInfoMateriel(MaterielFin, TobMateriel);

  //Chargement de la tobItem initiale avec les information de la tobMatériel de destination
  PR  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_COUT')));
  PA  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_PA')));
  PV  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_PV')));
  //
  TobEventMat.PutValue('BEM_PA',            PA);
  TobEventMat.PutValue('BEM_PV',            PV);
  TobEventMat.PutValue('BEM_PR',            PR);
  //
  TobEventMat.InsertOrUpdateDB;
  //
  FreeAndNil(TobMateriel);
  FreeAndNil(TobEventMat);
  //
end;

function MiseAJourParcMateriel(Item : TOB) : Boolean;
Var TobMateriel : TOB;
    FF          : String;
    PA          : Double;
    PR          : Double;
    PV          : Double;
    x           : Integer;
    STSQL       : string;
    Duree       : Double;
    NumEvent    : String;
    DateDeb     : TDateTime;
    DateFin     : TDateTime;
    HeureDeb    : TTime;
    HeureFin    : TTime;
    Materiel    : string;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
begin

	if not IsModifiable(Item) then
  begin
  	result := False;
    Exit;
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for x := 1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  NumEvent    := Item.GetValue('BPL_IDEVENT');

  Materiel    := Item.GetString('BPL_MATERIEL');
  TobMateriel := Tob.create('BTMATERIEL', nil, -1);
  if Materiel <> '' then ChargeInfoMateriel(Materiel, TobMateriel);

  //Chargement de la tobItem initiale avec les information de la tobMatériel de destination
  PR  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_COUT')));
  PA  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_PA')));
  PV  := StrToFloat(FormatFloat(FF, TobMateriel.GetVALUE('BMA_PV')));
  //
  StSQL :='UPDATE BTEVENTMAT SET ';
  StSQL := StSQL + 'BEM_AFFAIRE = "' + Item.GetValue('BPL_AFFAIRE') + '",';
  StSQL := StSQL + 'BEM_TIERS   = "' + Item.GetValue('BPL_TIERS')   + '",';
  StSQL := StSQL + 'BEM_DATEDEB = "' + USDateTime(StrToDateTime(Item.GetValue('BPL_DATEDEB'))) + '",';
  StSQL := StSQL + 'BEM_DATEFIN = "' + USDateTime(StrToDateTime(Item.GetValue('BPL_DATEFIN'))) + '",';
  StSQL := StSQL + 'BEM_HEUREDEB= "' + USDateTime(StrToDateTime(Item.GetValue('BPL_HEUREDEB'))) + '",';
  StSQL := StSQL + 'BEM_HEUREFIN= "' + USDateTime(StrToDateTime(Item.GetValue('BPL_HEUREFIN'))) + '",';
  StSQL := StSQL + 'BEM_NBHEURE=  "' + StrfPoint (Item.GetValue('BPL_DUREE'))    + '",';
  StSQL := StSQL + 'BEM_BTETAT=   "' + Item.GetValue('BPL_BTETAT') + '",';
  StSQL := StSQL + 'BEM_PA      =  ' + StrfPoint(PA) + ',';
  StSQL := StSQL + 'BEM_PR      =  ' + StrfPoint(PR) + ',';
  StSQL := StSQL + 'BEM_PV      =  ' + StrfPoint(PV) + ',';
  StSQL := StSQL + 'BEM_CODEMATERIEL = "'    + Materiel    + '" ';
  StSQL := StSQL + 'WHERE BEM_IDEVENTMAT = ' + NumEvent;

  if ExecuteSql (StSQL) <= 0 then Result := False;

  //libération de la tobMatériel
  FreeAndNil(TobMateriel);

end;


//
function CalculeDateEtireReduit (HeureFin : TTime; Cadencement : String): TTime;
Var  Delta       : Integer;
begin

  //Avant de faire quoi que ce soit il faut ajouter +1 à l'heure de fin
  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
    HeureFin := IncMinute (HeureFin, 15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
    HeureFin := IncMinute (HeureFin,30);
  end else if Cadencement = '003' then // gestion par heure
  begin
    HeureFin := IncHour (HeureFin, 1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > GetFinMatinee then
    begin
      Delta :=  MinutesBetween(heureFin,GetFinApresMidi);
    end else
    begin
      Delta := MinutesBetWeen(HeureFin,GetFinMatinee);
    end;
    HeureFin := incminute (HeureFin, Delta);
  end else if Cadencement = '005' then // gestion par Journée
  begin
    // pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
    {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;

  Result := HeureFin;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 07/04/2016
Modifié le ... :   /  /    
Description .. : Action possible sur un item Plan charge (Extension, 
Suite ........ : réduction,...)
Mots clefs ... : 
*****************************************************************}
function ReductionEtirementPlanCharge (Item : TOB; Cadencement : String) : boolean;
begin

  if not GestionDateItem(Item) then Exit;

   MajPlancharge (Item);

end;

function DeplacementPlanCharge (Item : TOB; Ressource : String; Cadencement : String) : boolean;
var CodeRessource   : String;
    Rep			 	      : String;
begin

  Result := True;
  //
  CodeRessource		  := Item.GetValue('BPL_RESSOURCE');
  //
  if not GestionDateItem(Item) then Exit;

  //Controle s'il n'y a pas de changement d'affectation
  if CodeRessource <> Ressource then
  Begin
    if PgiAsk('Confirmez-vous le changement d''affectation ?', rep)=mrNo then
    Begin
      Item.PutValue('BPL_RESSOURCE', Ressource);
    end;
  end;

  MAJPlanCharge(Item);;

end;

Function DuplicationPlanCharge(Item : TOB; Ressource : String; Cadencement : String) : boolean;
Var TobRessource : TOB;
    TobEventCha : TOB;
    STSQL       : string;
    QQ          : TQuery;
    FF          : String;
    x           : Integer;
    NewNumEvent : Integer;
    RessourceFin : String;
    NumEvent    : Integer;
begin

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for x := 1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  Result := false;

	if not IsModifiable(Item) then
  Begin
    PGIError('Cet événement n''est  pas modifiable', 'Duplication événement Parc/Matériel');
    exit;
  end;

  if not GetNumCompteur('BEC',iDate1900, NewNumEvent) then
  begin
    PGIError(TexteMessage[7], 'Duplication événement Chantier');
    exit;
  end;

  if not GestionDateItem(Item) then Exit;
  //
  RessourceFin := Item.GetValue('BPL_RESSOURCE');
  NumEvent     := Item.GetValue('BPL_IDEVENT');
  //
  StSQL := 'SELECT * FROM BTEVENTCHA WHERE BEC_IDEVENTCHA=' + IntToStr(NumEvent);
  QQ    := OpenSQL(StSql,False,1,'',True);
  if QQ.Eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  TobEventCha := TOB.Create('BTEVENTCHA',nil, -1);
  TobEventCha.SelectDB('BTEVENTCHA',QQ);
  //
  TobEventCha.PutValue('BEC_IDEVENTCHA',    NewNumEvent);
  TobEventCha.PutValue('BEC_RESSOURCE',     RessourceFin);
  //
  //
  TobEventCha.PutValue('BEC_DATEDEB',  Item.GetValue('BPL_DATEDEB')) ;
  TobEventCha.PutValue('BEC_DATEFIN',  Item.GetValue('BPL_DATEFIN')) ;
  TobEventCha.PutValue('BEC_HEUREDEB', Item.GetValue('BPL_HEUREDEB'));
  TobEventCha.PutValue('BEC_HEUREFIN', Item.GetValue('BPL_HEUREFIN'));
  TobEventCha.PutValue('BEC_NBHEURE',  Item.GetValue('BPL_DUREE'));

  TobRessource := Tob.create('RESSOURCE', nil, -1);
  if Ressource <> '' then ChargeInfoRessource(RessourceFin, '', TobRessource);
  //
  TobEventCha.InsertOrUpdateDB;
  //
  FreeAndNil(TobRessource);
  FreeAndNil(TobEventCha);

end;


Function MAJPlanCharge(Item : Tob) : Boolean;
Var TobRessource : TOB;
    FF          : String;
    x           : Integer;
    NumEvent    : Integer;
    STSQL       : string;
    Ressource   : String;
    Affaire     : String;
    Duree       : Double;
    DateDep     : TDateTime;
    Datefin     : TDateTime;
    HeureDeb    : TdateTime;
    HeureFin    : TDateTime;
begin

  Result := true;

	if not IsModifiable(Item) then
  begin
  	result := False;
    Exit;
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for x := 1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  NumEvent     := StrToInt(Item.GetValue('BPL_IDEVENT'));
  Affaire      := item.getValue('BPL_AFFAIRE');
  Ressource    := item.getValue('BPL_RESSOURCE');

  TobRessource := Tob.create('RESSOURCE', nil, -1);
  if Ressource <> '' then ChargeInfoRessource(Ressource,'', TobRessource);

  StSQL :='UPDATE BTEVENTCHA SET ';
  StSQL := StSQL + 'BEC_DATEDEB = "'    + USDateTime(StrToDateTime(Item.GetValue('BPL_DATEDEB'))) + '",';
  StSQL := StSQL + 'BEC_DATEFIN = "'    + USDateTime(StrToDateTime(Item.GetValue('BPL_DATEFIN'))) + '",';
  StSQL := StSQL + 'BEC_HEUREDEB= "'    + USDateTime(StrToDateTime(Item.GetValue('BPL_HEUREDEB'))) + '",';
  StSQL := StSQL + 'BEC_HEUREFIN= "'    + USDateTime(StrToDateTime(Item.GetValue('BPL_HEUREFIN'))) + '",';
  StSQL := StSQL + 'BEC_DUREE=  "'      + StrfPoint (Item.GetValue('BPL_DUREE')* 60)    + '",';
  StSQL := StSQL + 'BEC_FONCTION   = "' + TobRessource.GetValue('ARS_FONCTION1') + '", ';
  StSQL := StSQL + 'BEC_EQUIPERESS = "' + TobRessource.GetValue('ARS_EQUIPERESS')+ '", ';
  StSQL := StSQL + 'BEC_RESSOURCE  = "' + Ressource    + '",';
  stSQL := StSQL + 'BEC_AFFAIRE    = "' + Affaire      + '" ';
  StSQL := StSQL + 'WHERE BEC_IDEVENTCHA = ' + IntToStr(NumEvent);

  if ExecuteSql (StSQL) <= 0 then Result := False;

  //libération de la tobRessource
  FreeAndNil(TobRessource);             

end;

end.
