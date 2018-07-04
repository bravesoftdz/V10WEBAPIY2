unit SyntheseUtil;

interface

Uses Classes,HEnt1,windows,Paramsoc,stdctrls,ComCtrls,sysutils,HCtrls, Ent1,EntGC,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,DBGrids,HDB,
{$ENDIF}
     hmsgbox,LookUp,UTob,UTOM,FactUtil,FactComm,AGLInit,
     UtilPGI,HTB97,
     DicoBTP,Forms,SaisUtil,UtilRessource;

Function  SyntheseAffaireOLE(CodeEntete,ChampEntete,CodeDetail,ChampDetail,ChampCumule,UniteTemps ,Visa: string; DateEnc,DateDeb,DateFin : TDateTime):double;
Procedure ChargeTobActivite (TOBAct : TOB; CodeEntete,ChampEntete,CodeDetail,ChampDetail,ChampCumule ,ChampVisa: string; DateDeb,DateFin : TDateTime);
Procedure RegroupeTOBActivite( TOBGroup,TOBAct : TOB; ChampDetail,ChampCumule,Unite : string);


implementation

Function SyntheseAffaireOLE(CodeEntete,ChampEntete,CodeDetail,ChampDetail,ChampCumule,UniteTemps ,Visa: string; DateEnc,DateDeb,DateFin : TDateTime):double;
Var TOBAct,TOBGroup,TOBDet,TOBFind : TOB;
    ChargeTobAct : Boolean;
    Month,ii: integer;
    tmp1,ChampVisa:string;
BEGIN
Result := 0;
TOBGroup := nil;
if (CodeEntete = '') or (ChampEntete = '') or (CodeDetail = '') or (ChampDetail = '') then Exit;
if UniteTemps = '' then UniteTemps := VH_GC.AFMesureActivite;
if ChampCumule = '' then ChampCumule := 'ACT_QTE';
if (Visa <>'') then begin
   ChampVisa:= ' And (';
   ii:=0;
   tmp1 := ReadTokenSt(Visa);
   While tmp1  <>'' do
        begin
        if II <>0 then ChampVisa := ChampVisa + ' or ';
        champVisa:= ChampVisa + 'Act_etatVisa="'+ tmp1 +'"';
        inc (ii);
        tmp1 := ReadTokenSt(Visa);
        end;
   ChampVisa:= ChampVisa +')';
   end;

ChargeTobAct := true;
Datefin := FindeMois(DateFin);

if (TheTOB <> nil) And (TheTob.Detail.count <> 0) then
    BEGIN
    TOBDet:= TheTob.Detail[0];
    if TOBDet.GetValue(ChampEntete) = CodeEntete then
       BEGIN   // pas de changement d'affaire
       ChargeTobAct := false;
       TOBGroup := TheTob;
       END
    else //
       BEGIN
       TheTOB.Free; TheTOB := Nil;
       END;
    END;

if ChargeTobAct then
    BEGIN
    TobAct := TOB.Create ('lignes activites',Nil,-1);
    ChargeTobActivite (TOBAct, CodeEntete,ChampEntete,'',ChampDetail,ChampCumule,ChampVisa,DateDeb,DateFin);
    TOBGroup := TOB.Create('Activite regroupee',Nil,-1);
    RegroupeTOBActivite (TOBGroup,TOBAct,ChampDetail,ChampCumule,UniteTemps);
    TheTOB := TOBGroup;
    TOBAct.Free;
    END;

if TOBGroup <> Nil then
   BEGIN
   Month := GetPeriode(DateEnc);
   TOBFind:=TOBGroup.FindFirst([ChampDetail,'MOIS'],[CodeDetail,Month],False);
   if TOBFind <> Nil then Result := TOBFind.GetValue (ChampCumule);
   END;

END;

Procedure ChargeTobActivite (TOBAct : TOB; CodeEntete,ChampEntete,CodeDetail,ChampDetail,ChampCumule,ChampVisa : string; DateDeb,DateFin : TDateTime);
Var Q : TQuery;
    stSQL,stWhere : String;

BEGIN
stSQL := ''; stWhere := '';
stSQL := 'SELECT ' +ChampEntete + ', '+ ChampDetail +', ACT_DATEACTIVITE, ACT_UNITE, '+ ChampCumule +' FROM ACTIVITE ';
stWhere := ' WHERE ACT_TYPEARTICLE="PRE" ';
if CodeEntete <> '' then
    BEGIN
    stWhere := stWhere + 'AND ' + ChampEntete+'="'+ CodeEntete  + '" ';
    END;

if CodeDetail <> '' then
    BEGIN
    stWhere := stWhere +'AND '+ ChampDetail +'="'+ CodeDetail  + '" ';
    END;

if DateDeb <> iDate1900 then
    BEGIN
    stWhere := stWhere +'AND ' +' ACT_DATEACTIVITE >="'+UsDateTime(DateDeb) + '" ';
    END;

if DateFin <> iDate2099 then
    BEGIN
    stWhere := stWhere + 'AND ' +' ACT_DATEACTIVITE <="'+UsDateTime(DateFin) + '" ';
    END;

if CHampVIsa <>'' then stWhere:=StWhere + ChampVIsa;
Q := nil;
try
Q := OpenSQL(stSQL+stWhere,True,-1,'',true) ;
If Not Q.EOF then
    BEGIN
    TOBAct.LoadDetailDB ('synthese activite','','',Q,true);
    END;
Finally
 Ferme(Q);
 end;
END;

Procedure RegroupeTOBActivite( TOBGroup,TOBAct : TOB; ChampDetail,ChampCumule,Unite : string);
Var TOBDet,TOBActDet : TOB;
    i, Month: integer;
    d : double;
BEGIN
for i:=0 to TOBAct.Detail.count-1 do
    BEGIN
    TOBActDet := TOBAct.Detail[i];
    Month :=GetPeriode(TOBActDet.GetValue('ACT_DATEACTIVITE'));
    TOBDet:=TOBGroup.FindFirst([ChampDetail,'MOIS'],[TOBActDet.GetValue(ChampDetail),Month],False);
    if TOBDet <> Nil then
       BEGIN
       // MAJ D'un regroupement existant
       d := TOBActDet.GetValue(ChampCumule);
       if ChampCumule = 'ACT_QTE' then
          d := ConversionUnite( TOBActDet.GetValue('ACT_UNITE'),Unite,d);
       d:= TOBDet.GetValue(ChampCumule) + d;
       TOBDet.PutValue(ChampCumule,d);
       END
    else
       BEGIN
       // Création d'un nouveau regroupement
       TOBDet := TOB.Create('pas fixe',TOBGroup,-1);
       TobDet.Dupliquer (TOBActDet,False,True);
       TOBDet.AddChampSup('MOIS',False);
       TOBDet.PutValue ('MOIS',Month);
       d := TOBDet.GetValue(ChampCumule);
       if ChampCumule = 'ACT_QTE' then
          d := ConversionUnite( TOBDet.GetValue('ACT_UNITE'),Unite,d);
       TOBDet.PutValue(ChampCumule,d);
       END;
    END;
END;

end.
