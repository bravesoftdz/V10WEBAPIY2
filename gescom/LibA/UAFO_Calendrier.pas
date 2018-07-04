unit UAFO_Calendrier;

interface

uses  Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_main,
{$ENDIF}

{$IFDEF BTP}
	  CalcOleGenericBTP,
{$ENDIF}
      sysutils,
      HCtrls, HEnt1,UTob,ParamSoc,UtilPaieAffaire,HeureUtil,UAFO_Ferie;
(**********************************************************
Gestion des calendriers ressources et standards
PA le 05/11/99
*********************************************************)

// Classe des Horaires d'une journée.
type TAFO_Horaire = Class
   HDeb1,HFin1,HDeb2,HFin2,HDuree : double;
   TravailFerie,Ferie :Boolean;
   end;

// Classe calendrier
// Gestion des jours fériés, des calendriers (semaines, par dates ...)
Type TAFO_Calendrier = class
    Private
    FCode : String;
    FNomChamp  : String;
    FTypeCalendrier : string;
    FGestionFerie, FBase60 : Boolean;
    FTobCalenDate : Tob;
    FTobCalenSemaine : Tob;
    FTobDetailASupprimer : Tob;
    FJourFerie : TJourFerie;
    FDateDebut : TDateTime;
    FDateFin : TDateTime;
    Procedure ChargeCalendrierSemaine;

    Public
    gCodeStandard,gCodeRes,gSalarie:string;
    Constructor Create(TypeCalend, CodeStandard, CodeRessource,CodeSalarie : string;
                 Ferie:boolean;DateD,DateF:TDateTime; Base60 : Boolean = True); overload;
    Constructor Create(TypeCalend, CodeStandard, CodeRessource,CodeSalarie : string;
                 Ferie:boolean; Base60 : Boolean = True); reintroduce;overload;
    Destructor Destroy;override;
    function GetHoraire(DateJour :TDateTime):TAFO_Horaire;
    Function IsModifCalendrierDetail: Boolean;
    Procedure ChargeCalendrierDetail(DateD,DateF : TDateTime);
    Procedure SetCalendrierSemaine(HSemaine:TAFO_Horaire;Jour : Integer);
    Procedure SetCalendrierDetail(HDate:TAFO_Horaire; DateJour:TDateTime);
    Procedure AlimCleCalendrier(TobDet : Tob);
    Procedure UpdateCalendrierDetail;
    Procedure UpdateCalendrierSemaine;
    Procedure SetDateDebut (Valeur :TDateTime);
    Procedure SetDateFin (Valeur :TDateTime);
    Function  TestJourSemaine(DateJour:TDateTime;HDate:TAFO_Horaire):Boolean;
    Procedure DuplicationCalendrier (DateDebutRef,DateFinRef,DateDebutCible,DateFinCible:TDateTime);
    Procedure JourHeureTravail (var JourTravail : integer; var  HeureTravail : double);
    Function  PassageTobCalendrier100to60 (TobCalen : Tob ; Enter : Boolean) : Tob;
    Procedure ModifCleTobCalen(TypeCCible, CodeStdCible, ResCible,SalCible : String);

    property TypeCalendrier      : String    Read FTypeCalendrier;
    property TobCalenDate        : Tob       Read FTobCalenDate;
    property TobCalenSemaine     : Tob       Read FtobCalenSemaine;
    property TobDetailASupprimer : Tob       Read FTobDetailASupprimer Write FTobDetailASupprimer;
    property GestionFerie        : Boolean   Read FGestionFerie;
    property DateDebut           : TDateTime Read FDateDebut Write SetDateDebut;
    property DateFin             : TDateTime Read FDateFin Write setDateFin;
    property JourFerie           :TJourFerie Read FjourFerie;
    End;

// Fonctions indépendantes de traitement des calendriers
Function ExisteCalendrier (TypeC, CodeStd, Res,Sal : string) : Boolean;
Function DeleteCalendrier (TypeC, CodeStd, Res,Sal : string) : Boolean;
Procedure DuplicCalendrier (TypeCRef, CodeStdref, ResRef,SalRef,TypeCCible, CodeStdCible, ResCible,SalCible : string);

Procedure ChampCodeCalendrierTraite(TypeC, CodeStd, Res,Sal : string; Var Code , NomChamp : String);

Procedure AFLanceFiche_DuplicCalendrier;

implementation
Constructor TAFO_Calendrier.Create (TypeCalend, CodeStandard, CodeRessource,CodeSalarie : string;
                 Ferie:boolean; Base60 : Boolean = True);
Begin
gCodeStandard := CodeStandard;
gCoderes:= CodeRessource;
gSalarie:= CodeSalarie;
ChampCodeCalendrierTraite(TypeCalend, CodeStandard, CodeRessource,CodeSalarie, FCode, FNomChamp);

FTypeCalendrier := TypeCalend;
FGestionFerie := Ferie;
FBase60 := Base60;
SetDateDebut(iDate1900);SetDateFin(iDate1900);
ChargeCalendrierSemaine;
If (FGestionFerie) then FJourFerie:= TJourFerie.Create;
FTobDetailASupprimer := Nil;
End;

Constructor TAFO_Calendrier.Create (TypeCalend, CodeStandard, CodeRessource,CodeSalarie : string;
                 Ferie:boolean;DateD,DateF:TDateTime; Base60 : Boolean = True);
Begin
gCodeStandard := CodeStandard;
ChampCodeCalendrierTraite(TypeCalend, CodeStandard, CodeRessource,CodeSalarie, FCode, FNomChamp);

FTypeCalendrier := TypeCalend;
FGestionFerie := Ferie;
FBase60 := Base60;
ChargeCalendrierDetail(DateD,DateF);
ChargeCalendrierSemaine;
If (FGestionFerie) then FJourFerie:= TJourFerie.Create;
End;

destructor TAFO_Calendrier.Destroy;
begin
FTobCalenDate.Free;
FTobCalenSemaine.Free;
FJourFerie.Free; //FJourFerie:=nil;
inherited;
end;

Procedure TAFO_Calendrier.SetDateDebut (Valeur :TDateTime);
Begin
If ((Valeur <> 0) And(Valeur<>iDate1900)) Then FDateDebut:=Valeur Else FDateDebut:=iDate1900;
End;

Procedure TAFO_Calendrier.SetDateFin (Valeur :TDateTime);
Begin
If ((Valeur <> 0) And (Valeur<>iDate2099)) Then FDateFin:=Valeur else FDateFin:=iDate2099;
If (FDateFin < FDateDebut) Then FDatefin:= iDate2099;
End;

Procedure TAFO_Calendrier.ChargeCalendrierSemaine;
Var QQ :TQuery;
    stWhere : string;
Begin
FTobCalenSemaine:= Tob.Create('Liste Horaires',nil,-1);
If (TypeCalendrier ='STD') Then
   stWhere := ' AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"'
else
   stWhere := '';
QQ := nil;
Try
// PL le 06/03/02 : INDEX 2
// SELECT * : nombre de champs limité et nombre d'enregistrement restreint
 QQ := OpenSQL('SELECT * FROM CALENDRIER WHERE ACA_JOUR>=1 AND '+FNomChamp+'="'+FCode+'"'+ stWhere,true,-1,'',true);
 If Not QQ.EOF then FTobCalenSemaine.LoadDetailDB('CALENDRIER','','',QQ,True);
if (FBase60) then FTobCalenSemaine := PassageTobCalendrier100to60 ( FTobCalenSemaine , True );
Finally
 Ferme(QQ);
 End;
end;

Procedure TAFO_Calendrier.ChargeCalendrierDetail(DateD,DateF : TDateTime);
Var QQ : TQuery;
    stWhere : string;
Begin
If FTobCalenDate <> Nil Then BEGIN FTobCalenDate.Free; FTobCalenDate:= Nil; END;
SetDateDebut(DateD); SetDatefin(DateF);
If ((DateD=0) And (DateF=0)) Then Exit;

If (TypeCalendrier ='STD') Then
   stWhere := ' AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"'
else
   stWhere := '';

FTobCalenDate:=Tob.Create('Horaires detail',Nil,-1);
QQ := nil;
Try
// SELECT * : nombre de champs limité et nombre d'enregistrement restreint
QQ:=OpenSQL('SELECT * FROM CALENDRIER WHERE '+FNomChamp+'="'+FCode+
'"AND ACA_JOUR=0 AND ACA_DATE>="'+UsDateTime(DateDebut)+'" AND ACA_DATE<="'+UsDateTime(DateFin)+'"'+ stWhere,true,-1,'',true);
If Not QQ.EOF then FTobCalenDate.LoadDetailDB('CALENDRIER','','',QQ,True);
if (FBase60) then FTobCalenDate := PassageTobCalendrier100to60 ( FTobCalenDate , True );
Finally
   Ferme(QQ);
   TobCalenDate.SetAllModifie (FALSE);
   end;
End;

Function  TAFO_Calendrier.PassageTobCalendrier100to60 (TobCalen : Tob ; Enter : Boolean) : Tob;
Var Lig : Integer;
    TD  : TOB;
BEGIN
    for Lig:=0 to TobCalen.Detail.Count-1 do
        BEGIN
        TD:=TobCalen.Detail[lig];
        if (Enter) then
            BEGIN
            TD.PutValue('ACA_HEUREDEB1', BTPHeureBase100To60(TD.GetValue('ACA_HEUREDEB1')));
            TD.PutValue('ACA_HEUREFIN1', BTPHeureBase100To60(TD.GetValue('ACA_HEUREFIN1')));
            TD.PutValue('ACA_HEUREDEB2', BTPHeureBase100To60(TD.GetValue('ACA_HEUREDEB2')));
            TD.PutValue('ACA_HEUREFIN2', BTPHeureBase100To60(TD.GetValue('ACA_HEUREFIN2')));
            TD.PutValue('ACA_DUREE'    , BTPHeureBase100To60(TD.GetValue('ACA_DUREE'    )));
            END
        else // en sortie passage ne centième pour stockage
            BEGIN
            TD.PutValue('ACA_HEUREDEB1', HeureBase60To100(TD.GetValue('ACA_HEUREDEB1')));
            TD.PutValue('ACA_HEUREFIN1', BTPHeureBase100To60(TD.GetValue('ACA_HEUREFIN1')));
            TD.PutValue('ACA_HEUREDEB2', BTPHeureBase100To60(TD.GetValue('ACA_HEUREDEB2')));
            TD.PutValue('ACA_HEUREFIN2', BTPHeureBase100To60(TD.GetValue('ACA_HEUREFIN2')));
            TD.PutValue('ACA_DUREE'    , BTPHeureBase100To60(TD.GetValue('ACA_DUREE'    )));
            END;
        END;
Result := TobCalen;
END;


function TAFO_Calendrier.GetHoraire(DateJour :TDateTime):TAFO_Horaire;
Var Horaire: TAFO_Horaire;
    IYear,IMonth,IDay : word;
    Trouve : Boolean;
    i,Jour : Integer;
    TobDet : Tob;
Begin
Horaire:=TAFO_Horaire.Create;
Trouve := False;
DecodeDate(DateJour,IYear,IMonth,IDay);
// Test que le calendrier détail soit bien chargé sur cette période
// Sinon, on charge le mois en cours
If ((DateDebut > DateJour) Or (DateFin < DateJour)) Then
    ChargeCalendrierDetail(EncodeDate(IYear,IMonth,1),EncodeDate(IYear,IMonth,DaysPerMonth(IYear,IMonth)));
if GestionFerie then
    Begin
    Horaire.Ferie := FJourFerie.TestJourFerie(DateJour);
    End;

if (IDay>=1) and (IDay<=DaysPerMonth(IYear,IMonth)) then
    begin
    for i:=0 to TobCalenDate.Detail.Count-1 do
       BEGIN
       TobDet:=TobCalenDate.Detail[i];
       if (DateToStr(DateJour)=String(TobDet.GetValue('ACA_DATE'))) Then
          Begin
          Horaire.HDeb1:=double(TobDet.GetValue('ACA_HEUREDEB1'));
          Horaire.Hfin1:=double(TobDet.GetValue('ACA_HEUREFIN1'));
          Horaire.HDeb2:=double(TobDet.GetValue('ACA_HEUREDEB2'));
          Horaire.Hfin2:=double(TobDet.GetValue('ACA_HEUREFIN2'));
          Horaire.Hduree:=double(TobDet.GetValue('ACA_DUREE'));
          If (TobDet.GetValue('ACA_FERIETRAVAIL')='X') Then Horaire.TravailFerie:=True Else Horaire.TravailFerie:=False;
          Trouve := True;
          Break;
          End;
       End;

     If ((Not Trouve) And (Not Horaire.Ferie)) then
        Begin
        Jour:=DayOfWeek(DateJour)-1; if (Jour=0) Then Jour :=7;
        for i:=0 To TobCalenSemaine.Detail.Count-1 do
            Begin
            TobDet:=TobCalenSemaine.Detail[i];
            if (Jour = Integer(TobDet.GetValue('ACA_JOUR'))) Then
               Begin
               Horaire.HDeb1:=double(TobDet.GetValue('ACA_HEUREDEB1'));
               Horaire.Hfin1:=double(TobDet.GetValue('ACA_HEUREFIN1'));
               Horaire.HDeb2:=double(TobDet.GetValue('ACA_HEUREDEB2'));
               Horaire.Hfin2:=double(TobDet.GetValue('ACA_HEUREFIN2'));
               Horaire.Hduree:=double(TobDet.GetValue('ACA_DUREE'));
               Horaire.TravailFerie:=False;
               Trouve := True; Break;
               end;
            End;
        End;
     End;
If (Trouve And Horaire.Ferie And Not Horaire.TravailFerie) Then Trouve:=False;
If (Trouve) Then Result := Horaire Else Begin Result := Nil; Horaire.Free; end;
End;

Procedure TAFO_Calendrier.SetCalendrierSemaine(HSemaine:TAFO_Horaire;Jour : Integer);
Var  Trouve : Boolean;
     i : Integer;
     TobDet : TOB;
Begin
TobDet := nil;
Trouve:= False;
if TobCalenSemaine = nil then Exit;

// Recherche de l'enregistrement jour
for i:=0 to TobCalenSemaine.Detail.Count-1 do
    Begin
    TobDet:=TobCalenSemaine.Detail[i];
    If (Jour = integer(TobDet.GetValue('ACA_JOUR'))) Then begin Trouve:=True; break; end ;
    End;

If (Not Trouve) Then
    Begin
    // création d'un nouvel enregistrement
    TobDet :=Tob.Create('CALENDRIER',TobCalenSemaine,-1);
    AlimCleCalendrier(TobDet);
    TobDet.PutValue('ACA_JOUR',Jour);
    End;

if (TobDet = nil) then exit;
if (TobDet.GetValue('ACA_HEUREDEB1')<>HSemaine.HDeb1) then TobDet.PutValue('ACA_HEUREDEB1',HSemaine.HDeb1);
if (TobDet.GetValue('ACA_HEUREFIN1')<>HSemaine.HFin1) then TobDet.PutValue('ACA_HEUREFIN1',HSemaine.HFin1);
if (TobDet.GetValue('ACA_HEUREDEB2')<>HSemaine.HDeb2) then TobDet.PutValue('ACA_HEUREDEB2',HSemaine.HDeb2);
if (TobDet.GetValue('ACA_HEUREFIN2')<>HSemaine.HFin2) then TobDet.PutValue('ACA_HEUREFIN2',HSemaine.HFin2);
if (TobDet.GetValue('ACA_DUREE')    <>HSemaine.HDuree)then TobDet.PutValue('ACA_DUREE'    ,HSemaine.HDuree);
End;

Procedure TAFO_Calendrier.SetCalendrierDetail(HDate:TAFO_Horaire; DateJour:TDateTime);
Var  HeureBase,Trouve : Boolean;
     i : integer;
     TobDet : Tob;
Begin
TobDet := nil;
// Traitement de la TOB Détail
Trouve := False;
for i:=0 to TobCalenDate.Detail.Count-1 do
   BEGIN
   TobDet:=TobCalenDate.Detail[i];
   If (DateToStr(DateJour)=String(TobDet.GetValue('ACA_DATE'))) Then
      Begin Trouve :=True; break; End;
   End;
// Test pour savoir s'il s'agit de l'horaire de base.
HeureBase:=TestJourSemaine(DateJour,HDate);
  // MCD 11/09/03 si uniquement jour férié travaillé coché et horaire std, ne fct pas..
If (HeureBase) and (HDate.TravailFerie) then HeureBase:=False;


If ((Not Trouve) And (Not HeureBase))  Then
   Begin
   //Création d'un nouvel enregistrement
   TobDet := Tob.Create('CALENDRIER',TobCalenDate,-1);
   AlimCleCalendrier(TobDet);
   TobDet.PutValue('ACA_JOUR',0);
   TobDet.PutValue('ACA_DATE',DateJour);
   End;
If (Not HeureBase) and (TobDet <> nil) Then
    Begin
    TobDet.PutValue('ACA_HEUREDEB1',HDate.HDeb1); TobDet.PutValue('ACA_HEUREFIN1',HDate.HFin1);
    TobDet.PutValue('ACA_HEUREDEB2',HDate.HDeb2); TobDet.PutValue('ACA_HEUREFIN2',HDate.HFin2);
    TobDet.PutValue('ACA_DUREE',HDate.HDuree);
    If HDate.TravailFerie then TobDet.PutValue('ACA_FERIETRAVAIL','X') else TobDet.PutValue('ACA_FERIETRAVAIL','-');
    End;
If (Trouve And HeureBase) Then
    Begin
    If (TobDetailASupprimer = Nil) Then  TobDetailASupprimer:=Tob.Create('Horaires detail à supprimer',Nil,-1);
    if (TobDet <> nil) then
      TobDet.ChangeParent (TobDetailASupprimer,0);
    End;
End;

Procedure TAFO_Calendrier.AlimCleCalendrier(TobDet : Tob);
Var    Tmp, stdCal : string;
Begin
TobDet.PutValue(FNomChamp,FCode);
If (TypeCalendrier ='STD') Then
    Begin
    TobDet.PutValue('ACA_RESSOURCE','***');       // Zone de la clé non nulle
    TobDet.PutValue('ACA_SALARIE','***');       // Zone de la clé non nulle
    End
Else If (TypeCalendrier ='RES') Then
    Begin
    If gCodeStandard<>'' then stdCal:= gCodeStandard;
    TobDet.PutValue('ACA_STANDCALEN',stdCal);    // Zone de la clé non nulle
    Tmp := RechercheSalarieRessource(FCode,False);
    If (Tmp <>'') Then TobDet.PutValue('ACA_SALARIE',Tmp)
    Else TobDet.PutValue('ACA_SALARIE','***');
    End
Else If (TypeCalendrier ='SAL') Then
    Begin
    If gCodeStandard<>'' then stdCal:= gCodeStandard;
    TobDet.PutValue('ACA_STANDCALEN',StdCal);    // Zone de la clé non nulle
    Tmp := RechercheSalarieRessource(FCode,True);
    If (Tmp <>'') Then TobDet.PutValue('ACA_RESSOURCE',Tmp)
    Else TobDet.PutValue('ACA_RESSOURCE','***');
    End;
End;

Procedure TAFO_Calendrier.UpdateCalendrierDetail;
Begin
if (FBase60) then FTobCalenDate := PassageTobCalendrier100to60 ( TobCalenDate , False );
TobCalenDate.InsertOrUpdateDB;
TobCalenDate.SetAllModifie (FALSE);
If (TobDetailASupprimer <> Nil) Then TobDetailASupprimer.DeleteDB;
TobDetailASupprimer.Free; TobDetailASupprimer:= Nil;
End;

Function TAFO_Calendrier.IsModifCalendrierDetail: Boolean;
Begin
Result:=False;
if ((TobCalenDate.IsOneModifie) Or
(TobDetailASupprimer<>Nil) And (TobDetailASupprimer.Detail.Count >0))Then Result:=True;
End;

Procedure TAFO_Calendrier.UpdateCalendrierSemaine;
Begin
if (FBase60) then FTobCalenSemaine := PassageTobCalendrier100to60 ( TobCalenSemaine , False );
TobCalenSemaine.InsertOrUpdateDB;
TobCalenSemaine.SetAllModifie (FALSE);
End;

Function TAFO_Calendrier.TestJourSemaine(DateJour:TDateTime;HDate:TAFO_Horaire):Boolean;
Var JourSemaine, i : integer;
    Tobsemaine : Tob;
Begin
Result := False;
JourSemaine := DayOfWeek(Datejour)-1;
If (JourSemaine=0) Then JourSemaine :=7;
// Vérification que l'horaire différent de celui de base (personnalisé)
for i:=0 To TobCalenSemaine.Detail.Count-1 do
    Begin
    TobSemaine:=TobCalenSemaine.Detail[i];
    if (JourSemaine = Integer(TobSemaine.GetValue('ACA_JOUR'))) Then
       Begin
       If ((TobSemaine.GetValue('ACA_HEUREDEB1')= HDate.HDeb1)
       And (TobSemaine.GetValue('ACA_HEUREFIN1')= HDate.HFin1)
       And (TobSemaine.GetValue('ACA_HEUREDEB2')= HDate.HDeb2)
       And (TobSemaine.GetValue('ACA_HEUREFIN2')= HDate.HFin2)
       And (TobSemaine.GetValue('ACA_DUREE')=HDate.HDuree))Then Result:=True;
       Break;
       end;
    End;
End;

Procedure TAFO_Calendrier.DuplicationCalendrier (DateDebutRef,DateFinRef,DateDebutCible,DateFinCible:TDateTime);
Var TobCalendrierCible,TobDet : Tob;
    i,interval,decal : integer;
    QQ : TQuery;
    DateJour,DateJourCible : TDateTime;
    Horaire : TAFO_Horaire;
    stWhere : string;
Begin
// Suppression des enregistrements éventuels sur la période cible
If (TypeCalendrier ='STD') Then
   stWhere := ' AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"'
else
   stWhere := '';

TobCalendrierCible:=Tob.Create('Horaires detail cible',Nil,-1);
// PL le 06/03/02 : INDEX 2
// SELECT * : nombre de champs limité et nombre d'enregistrement restreint
QQ := OpenSQL('SELECT * FROM CALENDRIER WHERE ACA_JOUR=0 AND ACA_DATE>="'
    +UsDateTime(DateDebutCible)+'" AND ACA_DATE<="'+UsDateTime(DateFinCible)+'" AND '+FNomChamp+'="'+FCode+'"'+stWhere,true,-1,'',true);
If Not QQ.EOF then
    Begin
    TobCalendrierCible.LoadDetailDB('CALENDRIER','','',QQ,True);
    TobCalendrierCible.DeleteDB;
    End;
Ferme(QQ);
TobCalendrierCible.Free;

// Chargement de la tob détail sur la période de référence
ChargeCalendrierDetail(DateDebutRef,DateFinRef);

// Traitement du détail de la période de référence
Interval:=Trunc(DateFinRef-DateDebutRef)+1;
for i:=0 to TobCalenDate.Detail.Count-1 do
    BEGIN
    TobDet:=TobCalenDate.Detail[i];
    DateJour := StrToDate(TobDet.GetValue('ACA_DATE'));
    If ((DateJour>=DateDebutRef) or (DateJour<=DateFinRef)) Then
        Begin
        Horaire:=GetHoraire(DateJour);
        If (Not TestJourSemaine(DateJour,Horaire)) Then
            Begin
            // boucle sur toutes les dates à mettre à jour
            decal:=Trunc(DateJour-DateDebutRef);
            DateJourCible:=DateDebutCible+decal;
            While(DateJourCible < DateFinCible) do
                Begin
                SetCalendrierDetail(Horaire,DateJourCible);
                DateJourCible:=DateJourCible+interval;
                End;
            End;
        End;
    End;
UpdateCalendrierDetail;

End;

Procedure TAFO_Calendrier.JourHeureTravail (var JourTravail : integer; var  HeureTravail : double);
Var Horaire : TAFO_Horaire;
    LeJour  : TDateTime;
Begin
// Affectation des objets à chaque Jour du Grid.
LeJour := DateDebut;
JourTravail := 0; HeureTravail := 0;
While (LeJour <= DateFin) do
    BEGIN
    Horaire := GetHoraire(LeJour);
    if (Horaire <> Nil) Then
        BEGIN
        If Horaire.HDuree > 0 then
            BEGIN
            Inc(JourTravail);
            HeureTravail := HeureTravail + Horaire.HDuree;
            END;
        END;
    LeJour := LeJour+1;
    END;
END;

//******************************************************************************************
//************ Fonctions indépendantes de traitement des calendriers ***********************
//******************************************************************************************

Function ExisteCalendrier (TypeC, CodeStd, Res,Sal : string) : Boolean;
Var Code, NomChamp, stWhere : string;
BEGIN
ChampCodeCalendrierTraite(TypeC, CodeStd, Res,Sal, Code, NomChamp);
If (TypeC ='STD') Then
   stWhere := ' AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"'
else
   stWhere := ' AND ACA_STANDCALEN="'+CodeStd+'"';
// mcd 18/07/01 Result := ExisteSQL ('SELECT * FROM CALENDRIER WHERE '+ NomChamp+ '="'+ Code + '"');
Result := ExisteSQL ('SELECT '+NomChamp+' FROM CALENDRIER WHERE '+ NomChamp+ '="'+ Code +'"'+stWhere);
END;

Function DeleteCalendrier (TypeC, CodeStd, Res,Sal : string) : Boolean;
Var Code, NomChamp, pref : string;
BEGIN
ChampCodeCalendrierTraite(TypeC, CodeStd, Res,Sal, Code, NomChamp);
Pref := ReadTokenPipe(NomChamp,'_');
If NomChamp = 'STANDCALEN' then
  begin
   SupTablesLiees ('CALENDRIER',Pref+'_'+NomChamp, Code, 'AND ACA_RESSOURCE="***" AND ACA_SALARIE="***"' , True);
   Result :=SupTablesLiees ('CALENDRIERREGLE', 'ACG_'+NomChamp, Code, 'AND ACG_RESSOURCE="***" AND ACG_SALARIE="***"' , True);
  end
else
   Result :=SupTablesLiees ('CALENDRIER', Pref+'_'+NomChamp, Code,'', True);
END;


Procedure DuplicCalendrier (TypeCRef, CodeStdref, ResRef, SalRef, TypeCCible, CodeStdCible, ResCible, SalCible : string);
Var  Calendrier : TAFO_CALENDRIER;
     TobCalenSemaine, TobCalenDate : TOB;
BEGIN
  Calendrier := TAFO_Calendrier.Create (TypeCRef, CodeStdRef, ResRef, SalRef, False, iDate1900, iDate2099, False);
  TobCalenSemaine := Tob.create ('tob calendrier semaine', Nil, -1);
  TobCalenDate := Tob.create ('tob calendrier date', Nil, -1);
  try
    Calendrier.ModifCleTobCalen (TypeCCible, CodeStdCible, ResCible, SalCible);

    TobCalenSemaine.Dupliquer (Calendrier.FTobCalenSemaine, True, True, True);
    TobCalenDate.Dupliquer (Calendrier.FTobCalenDate, True, True, True);

    if Not TobCalenSemaine.InsertDB (Nil, False) then
      V_PGI.IoError := oeUnknown
    else
      if Not TobCalenDate.InsertDB (Nil, True) then
        V_PGI.IoError := oeUnknown;

  finally
    Calendrier.free;
    TobCalenSemaine.Free;
    TobCalenDate.Free;
  end;

END;

Procedure TAFO_Calendrier.ModifCleTobCalen(TypeCCible, CodeStdCible, ResCible,SalCible : String);
Var i : integer;
    TD: TOB;
BEGIN
FTypeCalendrier := TypeCCible;
ChampCodeCalendrierTraite(TypeCCible, CodeStdCible, ResCible,SalCible, FCode, FNomChamp);

for i := 0 to TobCalenSemaine.Detail.Count-1 do
    BEGIN
    TD := TobCalenSemaine.Detail[i];
    AlimCleCalendrier (TD);
    END;

for i := 0 to TobCalenDate.Detail.Count-1 do
    BEGIN
    TD := TobCalenDate.Detail[i];
    AlimCleCalendrier (TD);
    END;
END;

Procedure AFLanceFiche_DuplicCalendrier;
begin
AGLLanceFiche ('AFF','DUPLICCALENDRIER','','','');
end;

Procedure ChampCodeCalendrierTraite(TypeC, CodeStd, Res,Sal : string; Var Code , NomChamp : String);
BEGIN
If (TypeC='STD') Then Begin Code := CodeStd; NomChamp :=  'ACA_STANDCALEN'; End Else
If (TypeC='RES') Then Begin Code := Res;     NomChamp :=  'ACA_RESSOURCE';  End Else
If (TypeC='SAL') Then Begin Code := Sal;     NomChamp :=  'ACA_SALARIE';    End;
END;

end.
