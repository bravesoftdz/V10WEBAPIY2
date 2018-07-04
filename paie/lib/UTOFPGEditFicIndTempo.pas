unit UTOFPGEditFicIndTempo;

interface
uses StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,UTOB,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,PGIenv,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,ParamSoc,HQry,HTB97;


Type
     TOF_PGFICHEINDTEMPO = Class (TOF)
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; override;
       procedure GenereEnrFicheInd ( Sender : TObject);
       procedure GenereEnrCumul ( Sender : TObject);
       Function  RendSqlHistorique(ChampSum : String; TabPeriodeDebut,TabPeriodeFin : Array of TDateTime) : String;
     END ;

implementation

{ TOF_PGFICHEINDTEMPO }

procedure TOF_PGFICHEINDTEMPO.OnArgument(Arguments: String);
Var
Btn : TToolBarButton97;
begin
  inherited;
Btn := TToolBarButton97(GetControl('BGENERE'));
if Btn<>nil then Btn.OnClick:=GenereEnrFicheInd;
Btn := TToolBarButton97(GetControl('BCUMUL'));
if Btn<>nil then Btn.OnClick:=GenereEnrCumul;
SetControltext('PT1_USER',V_PGI.User);

end;
procedure TOF_PGFICHEINDTEMPO.OnUpdate;
Begin
  inherited;

End;


procedure TOF_PGFICHEINDTEMPO.GenereEnrFicheInd ( Sender : TObject);
Var
  DateDebut,DateFin,DateUtil    : TDateTime;
  StDelete,StRub,StCumul,St,Sti,RuptSal,RuptEtab : String;
  TabPeriodeDebut,TabPeriodeFin : Array [1..12 ]of TDateTime;
  NbPeriode,i                   : Integer;
begin
  NbPeriode := 0;
  St        := '';
{ Traitement des critères }
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin   := StrToDate(GetControlText('DATEFIN'));
  if GetControlText('CSAL')='X' then RuptSal:='PMAIN.PHB_SALARIE,' else RuptSal:='"***",';
  if GetControlText('CETAB')='X' then RuptEtab:='PMAIN.PHB_ETABLISSEMENT,' else RuptEtab:='"***",';
  { Constitution des sessions de paie }
  DateUtil:=DateDebut;
  While DateUtil<DateFin do
    Begin
    Inc(NbPeriode);
    TabPeriodeDebut[NbPeriode] := DateUtil;
    if FinDeMois(DateUtil)<DateFin then TabPeriodeFin[NbPeriode] := FinDeMois(DateUtil)
    else                                TabPeriodeFin[NbPeriode] := DateFin;
    DateUtil :=FinDeMois(DateUtil)+1;
    if NbPeriode=12 then break;
    End;
{ Construction des requêtes }
  { Suppression des données existantes }
  StDelete :=  'DELETE FROM PTMPPAYE1 WHERE PT1_USER="'+V_PGI.User+'" '+
               'AND PT1_DATEDEBUT>="'+UsDateTime(DateDebut)+'" '+
               'AND PT1_DATEFIN<="'+UsDateTime(DateFin)+'"';

  { Insertion Historique de paie}
  StRub :=  'INSERT INTO PTMPPAYE1 SELECT "'+V_PGI.User+'","'+UsDateTime(DateDebut)+'",'+
               '"'+UsDateTime(DateFin)+'",'+RuptEtab+RuptSal+
               'PMAIN.PHB_NATURERUB,PMAIN.PHB_ORDREETAT,PMAIN.PHB_RUBRIQUE,PHB_LIBELLE';
  // Montant remunération ou salarial
    St := RendSqlHistorique('PHB_MTREM+PHB_MTSALARIAL',TabPeriodeDebut,TabPeriodeFin);
  // Montant Base de cotisation
    St := St + RendSqlHistorique('PHB_BASECOT',TabPeriodeDebut,TabPeriodeFin);
  // Montant patronal
    St := St + RendSqlHistorique('PHB_MTPATRONAL',TabPeriodeDebut,TabPeriodeFin);
  if GetControlText('CSAL')='X' then RuptSal:='PMAIN.PHB_SALARIE,' else RuptSal:='';
  if GetControlText('CETAB')='X' then RuptEtab:='PMAIN.PHB_ETABLISSEMENT,' else RuptEtab:='';
  StRub := StRub + St +
              ' FROM HISTOBULLETIN PMAIN '+
              'WHERE PMAIN.PHB_DATEDEBUT>="'+UsDateTime(DateDebut)+'" '+
              'AND PMAIN.PHB_DATEFIN<="'+UsDateTime(DateFin)+'" '+
              'AND (((PMAIN.PHB_MTREM<>0 OR PMAIN.PHB_MTSALARIAL<>0 OR PMAIN.PHB_MTPATRONAL<>0) '+
              'AND PMAIN.PHB_IMPRIMABLE="X") AND PMAIN.PHB_NATURERUB<>"BAS") '+
              'GROUP BY '+RuptEtab+RuptSal+
              'PMAIN.PHB_NATURERUB,PMAIN.PHB_ORDREETAT,PMAIN.PHB_RUBRIQUE,PMAIN.PHB_LIBELLE';

  // Les cumuls de paie
  if GetControlText('CSAL')='X' then RuptSal:='PMAIN.PHC_SALARIE,' else RuptSal:='"***",';
  if GetControlText('CETAB')='X' then RuptEtab:='PMAIN.PHC_ETABLISSEMENT,' else RuptEtab:='"***",';
  StCumul :='INSERT INTO PTMPPAYE1 SELECT "'+V_PGI.User+'","'+UsDateTime(DateDebut)+'",'+
               '"'+UsDateTime(DateFin)+'",'+RuptEtab+RuptSal+
               '"BUL",3,PMAIN.PHC_CUMULPAIE,""';
  For i:=1 to 12 do
    Begin
    Sti:=IntToStr(i);
    if TabPeriodeDebut[i]>idate1900 then
      StCumul := StCumul+',(SELECT SUM(P'+Sti+'.PHC_MONTANT) FROM HISTOCUMSAL P'+Sti+' '+
               'WHERE P'+Sti+'.PHC_DATEDEBUT>="'+USDateTime(TabPeriodeDebut[i])+'" '+
               'AND P'+Sti+'.PHC_DATEFIN<="'+USDateTime(TabPeriodeFin[i])+'" '+
               'AND P'+Sti+'.PHC_SALARIE=PMAIN.PHC_SALARIE '+
               'AND P'+Sti+'.PHC_CUMULPAIE=PMAIN.PHC_CUMULPAIE )'
    else
      StCumul := StCumul+',0';
    End;

  if GetControlText('CSAL')='X' then RuptSal:='PMAIN.PHC_SALARIE,' else RuptSal:='';
  if GetControlText('CETAB')='X' then RuptEtab:='PMAIN.PHC_ETABLISSEMENT,' else RuptEtab:='';
  StCumul := StCumul  + ',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 '+
              'FROM HISTOCUMSAL PMAIN '+
              'WHERE PMAIN.PHC_DATEDEBUT>="'+UsDateTime(DateDebut)+'" '+
              'AND PMAIN.PHC_DATEFIN<="'+UsDateTime(DateFin)+'" '+
              'AND PMAIN.PHC_REPRISE="-" '+
              'AND (PMAIN.PHC_CUMULPAIE="01" OR PMAIN.PHC_CUMULPAIE="02" '+
              'OR PMAIN.PHC_CUMULPAIE="08" OR PMAIN.PHC_CUMULPAIE="09" '+
              'OR PMAIN.PHC_CUMULPAIE="10" OR PMAIN.PHC_CUMULPAIE="20" '+
              'OR PMAIN.PHC_CUMULPAIE="30")'+
              'GROUP BY '+RuptEtab+RuptSal+'PMAIN.PHC_CUMULPAIE';
{ Execution des requêtes }
  ExecuteSql(StDelete);
  ExecuteSql(StRub);
  ExecuteSql(StCumul);
  ExecuteSql( 'UPDATE PTMPPAYE1 SET PT1_NATURERUB="CUM" '+
              'WHERE PT1_USER="'+V_PGI.User+'" AND PT1_DATEDEBUT="'+UsDateTime(DateDebut)+'" '+
              'AND PT1_DATEFIN="'+UsDateTime(DateFin)+'" AND PT1_RUBRIQUE="08"');
  ExecuteSql( 'UPDATE PTMPPAYE1 SET PT1_ORDREETAT=8 '+
              'WHERE PT1_USER="'+V_PGI.User+'" AND PT1_DATEDEBUT="'+UsDateTime(DateDebut)+'" '+
              'AND PT1_DATEFIN="'+UsDateTime(DateFin)+'" AND (PT1_RUBRIQUE="09" OR PT1_RUBRIQUE="10")');


PgiInfo('Traitement terminé',Ecran.caption);
End;


Function  TOF_PGFICHEINDTEMPO.RendSqlHistorique(ChampSum : String; TabPeriodeDebut,TabPeriodeFin : Array of TDateTime) : String;
Var
  Sti : String;
  i   : Integer;
begin
  For i:=1 to 12 do
    Begin
    Sti:=IntToStr(i);
    if TabPeriodeDebut[i-1]>idate1900 then
       Begin
       if ChampSum='PHC_MONTANT' then
          Result := Result+',(SELECT SUM(P'+Sti+'.'+ChampSum+') FROM HISTOCUMSAL P'+Sti+' '+
                 'WHERE P'+Sti+'.PHC_DATEDEBUT>="'+USDateTime(TabPeriodeDebut[i-1])+'" '+
                 'AND P'+Sti+'.PHC_DATEFIN<="'+USDateTime(TabPeriodeFin[i-1])+'" '+
                 'AND P'+Sti+'.PHC_SALARIE=PMAIN.PHC_SALARIE '+
                 'AND P'+Sti+'.PHC_CUMULPAIE=PMAIN.PHC_CUMULPAIE )'
       else
          Result := Result+',(SELECT SUM(P'+Sti+'.'+ChampSum+') FROM HISTOBULLETIN P'+Sti+' '+
                 'WHERE P'+Sti+'.PHB_DATEDEBUT>="'+USDateTime(TabPeriodeDebut[i-1])+'" '+
                 'AND P'+Sti+'.PHB_DATEFIN<="'+USDateTime(TabPeriodeFin[i-1])+'" '+
                 'AND P'+Sti+'.PHB_SALARIE=PMAIN.PHB_SALARIE '+
                 'AND P'+Sti+'.PHB_RUBRIQUE=PMAIN.PHB_RUBRIQUE '+
                 'AND P'+Sti+'.PHB_NATURERUB=PMAIN.PHB_NATURERUB )';
       End
    else
      Result := Result+',0';
    End;
end;

procedure TOF_PGFICHEINDTEMPO.GenereEnrCumul(Sender: TObject);
Var
  DateDebut,DateFin,DateUtil    : TDateTime;
  StDelete,StCumul,St           : String;
  TabPeriodeDebut,TabPeriodeFin : Array [1..12 ]of TDateTime;
  NbPeriode                     : Integer;
begin
  NbPeriode := 0;
  St        := '';
{ Traitement des critères }
  DateDebut := StrToDate(GetControlText('DATEDEBUT'));
  DateFin   := StrToDate(GetControlText('DATEFIN'));
  { Constitution des sessions de paie }
  DateUtil:=DateDebut;
  While DateUtil<DateFin do
    Begin
    Inc(NbPeriode);
    TabPeriodeDebut[NbPeriode] := DateUtil;
    if FinDeMois(DateUtil)<DateFin then TabPeriodeFin[NbPeriode] := FinDeMois(DateUtil)
    else                                TabPeriodeFin[NbPeriode] := DateFin;
    DateUtil :=FinDeMois(DateUtil)+1;
    if NbPeriode=12 then break;
    End;
{ Construction des requêtes }
  { Suppression des données existantes }
  StDelete :=  'DELETE FROM PTMPPAYE1 WHERE PT1_USER="'+V_PGI.User+'" '+
               'AND PT1_DATEDEBUT>="'+UsDateTime(DateDebut)+'" '+
               'AND PT1_DATEFIN<="'+UsDateTime(DateFin)+'"';

  { Insertion cumuls de paie}
  StCumul :=  'INSERT INTO PTMPPAYE1 SELECT "'+V_PGI.User+'","'+UsDateTime(DateDebut)+'",'+
               '"'+UsDateTime(DateFin)+'",PMAIN.PHC_ETABLISSEMENT,PMAIN.PHC_SALARIE,'+
               '"ZZZ",0,PMAIN.PHC_CUMULPAIE,"test"';

  // Montant Cumul
    St := RendSqlHistorique('PHC_MONTANT',TabPeriodeDebut,TabPeriodeFin);
    St := St  + ',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 ';



  StCumul := StCumul + St +
              ' FROM HISTOCUMSAL PMAIN '+
              'WHERE PMAIN.PHC_DATEDEBUT>="'+UsDateTime(DateDebut)+'" '+
              'AND PMAIN.PHC_DATEFIN<="'+UsDateTime(DateFin)+'" '+
              'GROUP BY PMAIN.PHC_ETABLISSEMENT,PMAIN.PHC_SALARIE,PMAIN.PHC_CUMULPAIE';

{ Execution des requêtes }
  ExecuteSql(StDelete);
  ExecuteSql(StCumul);
PgiInfo('Traitement terminé',Ecran.caption);
end;

Initialization
registerclasses([TOF_PGFICHEINDTEMPO]);
end.
