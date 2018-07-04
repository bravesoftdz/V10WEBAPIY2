unit BPEclatement;

interface

uses Sysutils,Classes,
     {$IFDEF EAGLCLIENT}
     UTOB,
     {$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     HCtrls,HEnt1;

procedure CalculEclatementLoi(const codeSession,TypeEclat:hString);
function ChercheNomValeurAxeDelai(const maille:hString;
         i:integer;DateDeb:TDateTime):TDateTime;


procedure InitialisationPrctDelaiParRapportHisto(const mode,codeSession,codeRestriction,codeJoinR:hString;
          dateDebRef,dateFinRef:TDateTime;
          var TabQteLoi:array of double);
procedure InitialisationPrctDelaiParRapportHistoPgi(const mode,codeSession,codeRestriction,codeJoinR:hString;
          dateDebRef,dateFinRef:TDateTime;
          var TabQteLoi:array of double);

procedure InitialisationPrctTailleParRapportHisto(const CodeTaille,codeRestriction:hString;
          dateDeb,dateFin:TDateTime;
          var TabQteLoi:array of double);


implementation

uses CstCommun,Uutil,BPBasic,BPFctSession,BPFctArbre,UDatamem,BPUtil,StrUtils,ed_tools,HmsgBox

;

function ChercheNomValeurAxeDelai(const maille:hString;
         i:integer;DateDeb:TDateTime):TDateTime;
var DelaiMaille:hString;
    code:TDateTime;
    code1, AnSemaine:hString;
    An,Mois,Jour:word;
    PremJDep,DateDebCMi,DateFinCMi,DateI:TDateTime;
    MoisI:integer;
    NumT,AnT,NbAn:integer;
    anDeb,moisDeb,QuinzDeb,NumQuinzDeb:word;
    Per:TperiodeP;
begin
 code:=0;
 DelaiMaille:=DonneParamS(ps_BPDelaiMaille);
 decodeDate(DateDeb,An,Mois,Jour);
 case maille[1] of
  '1':begin
       code:=PLUSDATE(DateDeb,i,'J');
      end;
  '2':begin
      //numéro de semaine
      DateI := DateDeb + (i*7);
      AnSemaine := DonneAnSemaine(DateI);
      PremJDep:=PremierJourSemaine(ValeurI(Copy(AnSemaine,4,2)), an2_an4(ValeurI(Copy(AnSemaine,1,2))));
       //resultat
       code:=PremJDep;
      end;
  '3':begin
       Date_AnNumQuinz(DateDeb,anDeb,moisDeb,QuinzDeb,NumQuinzDeb);
       NumQuinzDeb:=NumQuinzDeb+i;
       { EVI / Correction : "Mauvais encodage de la date" avec NumQuinzDeb>24 }
       NbAn := Trunc(NumQuinzDeb/24);
       if (NumQuinzDeb mod 24)=0 then NbAn := NbAn-1;
       NumQuinzDeb:=NumQuinzDeb-(NbAn*24);
       anDeb:=anDeb+NbAn;
       code:=AnNumQuinz_Date(anDeb,NumQuinzDeb);
      end;
  '4':begin
       MoisI:=Mois+i;
       { EVI / Correction : "Mauvais encodage de la date" }
       NbAn := Trunc(MoisI/12);
       //if Frac(MoisI/12)=0 then NbAn := NbAn-1;
       if (MoisI mod 12)=0 then NbAn := NbAn-1;
       MoisI:=MoisI-(NbAn*12);
       An:=An+NbAn;
       case DelaiMaille[1] of
        '1':code1:='01';
        '2':code1:='15';
        '3':begin
             case MoisI of
              1,3,5,7,8,10,12:code1:='31';
              2:code1:='28';
              else code1:='30';
             end;
            end;
       end;
       code:=EncodeDate(An,MoisI,VALEURI(code1));
      end;
  '5':begin
       Per:=AddPeriodeP(DateDeb,i);
       code:=Per.datedeb;
      end;
  '6':begin
       DonneNumTrimestre(DateDeb,NumT,AnT);
       NumT:=NumT+i;
       if NumT>4
        then
         begin
          NumT:=4-NumT;
          AnT:=AnT+1;
         end;
       DonneDateDebFinTrimestre(NumT,AnT,DateDebCMi,DateFinCMi);
       code:=DateDebCMi;
      end;
  '7':begin  
       DonneNumQuadrimestre(DateDeb,NumT,AnT);
       NumT:=NumT+i;
       if NumT>3
        then
         begin
          NumT:=3-NumT;
          AnT:=AnT+1;
         end;
       DonneDateDebFinQuadrimestre(NumT,AnT,DateDebCMi,DateFinCMi);
       code:=DateDebCMi;  
      end;
 end;
 result:=code;
end;


procedure RemplitTabPrctEclatTaille(const Loi:hString;
          var tabValeurPrct:array of double;
          var DateDebLoi:TDateTime;
          var LoiGrilleTaille:hString);
var Q:TQuery;
    i:integer;
begin
 Q:=MOpenSql('SELECT QBL_PRCTTAILLE1,QBL_PRCTTAILLE2,QBL_PRCTTAILLE3,'+
             'QBL_PRCTTAILLE4,QBL_PRCTTAILLE5,QBL_PRCTTAILLE6,'+
             'QBL_PRCTTAILLE7,QBL_PRCTTAILLE8,QBL_PRCTTAILLE9,'+
             'QBL_PRCTTAILLE10,QBL_PRCTTAILLE11,QBL_PRCTTAILLE12,'+
             'QBL_PRCTTAILLE13,QBL_PRCTTAILLE14,QBL_PRCTTAILLE15,'+
             'QBL_PRCTTAILLE16,QBL_PRCTTAILLE17,QBL_PRCTTAILLE18,'+
             'QBL_PRCTTAILLE19,QBL_PRCTTAILLE20,QBL_DATEDEBTAI,QBL_CODETAILLE '+
             ' FROM QBPLOITAILLE '+
             'WHERE QBL_CODELOITAILLE="'+Loi+'"',
             'BPEclatement (RemplitTabPrctEclatTaille).',true);
 if not Q.eof
  then
   begin
    for i:=0 to 19 do
     tabValeurPrct[i]:=Q.Fields[i].asFloat/100;

    DateDebLoi:=Q.fields[20].AsDateTime;
    LoiGrilleTaille:=Q.fields[21].AsString;
   end;
 ferme(Q);
end;

procedure RemplitTabPrctEclat(const codeSession,Loi:hString;
          var maille,mode:hString;
          var tabValeurPrct:array of double;
          var DateDebLoi,DateFinLoi:TDateTime);
var Q:TQuery;
    i:integer;
    codeWhere:hString;
begin

 //cas pgi => code session
 codeWhere:='';
 if not BPOkOrli
  then codeWhere:=' AND QBO_CODESESSION="'+codeSession+'" ';

 Q:=MOpenSql('SELECT QBO_MAILLE,QBO_MODEECLAT,'+
             'QBO_BPM1,QBO_BPM2,QBO_BPM3,QBO_BPM4,QBO_BPM5,'+
             'QBO_BPM6,QBO_BPM7,QBO_BPM8,QBO_BPM9,QBO_BPM10,'+
             'QBO_BPM11,QBO_BPM12,QBO_BPM13,QBO_BPM14,QBO_BPM15,'+
             'QBO_BPM16,QBO_BPM17,QBO_BPM18,QBO_BPM19,QBO_BPM20,'+
             'QBO_BPM21,QBO_BPM22,QBO_BPM23,QBO_BPM24,QBO_BPM25,'+
             'QBO_BPM26,QBO_BPM27,QBO_BPM28,QBO_BPM29,QBO_BPM30,'+
             'QBO_BPM31,QBO_BPM32,QBO_BPM33,QBO_BPM34,QBO_BPM35,'+
             'QBO_BPM36,QBO_BPM37,QBO_BPM38,QBO_BPM39,QBO_BPM40,'+
             'QBO_BPM41,QBO_BPM42,QBO_BPM43,QBO_BPM44,QBO_BPM45,'+
             'QBO_BPM46,QBO_BPM47,QBO_BPM48,QBO_BPM49,QBO_BPM50,'+
             'QBO_BPM51,QBO_BPM52,QBO_DATEDEBECLAT,QBO_DATEFINECLAT '+
             ' FROM QBPLOI '+
             'WHERE QBO_CODELOI="'+Loi+'" '+codeWhere,
             'BPEclatement (RemplitTabPrctEclat).',true);
 if not Q.eof
  then
   begin
    maille:=Q.Fields[0].asString;
    mode:=Q.Fields[1].asString;
    for i:=0 to 51 do
     tabValeurPrct[i]:=Q.Fields[i+2].asFloat/100;

    DateDebLoi:=Q.fields[54].AsDateTime;   
    DateFinLoi:=Q.fields[55].AsDateTime;
   end;
 ferme(Q);
end;


function RemplitTabPrctEclatLoi(codeSession,GrilleTaille:hString;OkTaille:boolean;
         TabValAxe:array of hString;
         TStCodeAxe:TStringList;
         var maille,LoiGrilleTaille:hString;
         var tabValeurPrct:array of double;
         var DateDebLoi,DateFinLoi:TDateTime):boolean;
var Q:TQuery;
    loi,mode,codeWhere:hString;
    codeaxeR1,valaxeR1,codeaxeR2,valaxeR2,codeaxeR3,valaxeR3:hString;
    codeaxeR4,valaxeR4,codeaxeR5,valaxeR5,codeaxeR6,valaxeR6:hString;
    codeaxeR7,valaxeR7,codeaxeR8,valaxeR8,codeaxeR9,valaxeR9:hString;
    codeaxeR10,valaxeR10:hString;
    trouve:integer;
    trouveLoi:hString;
    codeSql:hString;
begin
 trouve:=-1;
 trouveLoi:='';
 result:=false;
 //cas pgi => code session
 codeWhere:='';
 if not BPOkOrli
  then codeWhere:=' WHERE QBO_CODESESSION="'+codeSession+'" ';
 //cherche la meilleur loi
 //codesql
 if OkTaille
  then codeSql:='SELECT QBL_CODELOITAILLE,'+
             'QBL_CODEAXELT1,QBL_VALAXELT1,QBL_CODEAXELT2,QBL_VALAXELT2,'+
             'QBL_CODEAXELT3,QBL_VALAXELT3,QBL_CODEAXELT4,QBL_VALAXELT4,'+
             'QBL_CODEAXELT5,QBL_VALAXELT5,QBL_CODEAXELT6,QBL_VALAXELT6,'+
             'QBL_CODEAXELT7,QBL_VALAXELT7,QBL_CODEAXELT8,QBL_VALAXELT8,'+
             'QBL_CODEAXELT9,QBL_VALAXELT9,QBL_CODEAXELT10,QBL_VALAXELT10 '+
             ' FROM QBPLOITAILLE WHERE QBL_CODETAILLE="'+GrilleTaille+'"'
  else codeSql:='SELECT QBO_CODELOI,'+
             'QBO_CODEAXER1,QBO_VALAXER1,QBO_CODEAXER2,QBO_VALAXER2,'+
             'QBO_CODEAXER3,QBO_VALAXER3,QBO_CODEAXER4,QBO_VALAXER4,'+
             'QBO_CODEAXER5,QBO_VALAXER5,QBO_CODEAXER6,QBO_VALAXER6,'+
             'QBO_CODEAXER7,QBO_VALAXER7,QBO_CODEAXER8,QBO_VALAXER8,'+
             'QBO_CODEAXER9,QBO_VALAXER9,QBO_CODEAXER10,QBO_VALAXER10 '+
             ' FROM QBPLOI '+codeWhere;

 Q:=MOpenSql(codeSql,'BPEclatement (RemplitTabPrctEclat).',true);
 while not Q.eof do
  begin
   loi:=Q.Fields[0].asString;
   codeaxeR1:=Q.Fields[1].asString;
   valaxeR1:=Q.Fields[2].asString;
   codeaxeR2:=Q.Fields[3].asString;
   valaxeR2:=Q.Fields[4].asString;
   codeaxeR3:=Q.Fields[5].asString;
   valaxeR3:=Q.Fields[6].asString;
   codeaxeR4:=Q.Fields[7].asString;
   valaxeR4:=Q.Fields[8].asString;
   codeaxeR5:=Q.Fields[9].asString;
   valaxeR5:=Q.Fields[10].asString;
   codeaxeR6:=Q.Fields[11].asString;
   valaxeR6:=Q.Fields[12].asString;
   codeaxeR7:=Q.Fields[13].asString;
   valaxeR7:=Q.Fields[14].asString;
   codeaxeR8:=Q.Fields[15].asString;
   valaxeR8:=Q.Fields[16].asString;
   codeaxeR9:=Q.Fields[17].asString;
   valaxeR9:=Q.Fields[18].asString;
   codeaxeR10:=Q.Fields[19].asString;
   valaxeR10:=Q.Fields[20].asString;

   //on tombe sur une loi particuliere
   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR2)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR3)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR4)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR5)<>0)
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)],valaxeR6)<>0)
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)],valaxeR7)<>0)
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)],valaxeR8)<>0)
      and (codeaxeR9<>'') and (TStCodeAxe.IndexOf(codeaxeR9)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR9)],valaxeR9)<>0)
      and (codeaxeR10<>'') and (TStCodeAxe.IndexOf(codeaxeR10)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR10)],valaxeR10)<>0)
      and (trouve<10)
    then
     begin
      trouve:=10;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR1)<>0)
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)],valaxeR1)<>0)
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)],valaxeR1)<>0)
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)],valaxeR1)<>0)
      and (codeaxeR9<>'') and (TStCodeAxe.IndexOf(codeaxeR9)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR9)],valaxeR1)<>0)
      and (codeaxeR10='')
      and (trouve<9)
    then
     begin
      trouve:=9;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR1)<>0)
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)],valaxeR1)<>0)
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)],valaxeR1)<>0)
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)],valaxeR1)<>0)
      and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<8)
    then
     begin
      trouve:=8;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR1)<>0)
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)],valaxeR1)<>0)
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)],valaxeR1)<>0)
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<7)
    then
     begin
      trouve:=7;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR1)<>0)
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)],valaxeR1)<>0)
      and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<6)
    then
     begin
      trouve:=6;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)],valaxeR1)<>0)
      and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<5)
    then
     begin
      trouve:=5;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)],valaxeR1)<>0)
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<4)
    then
     begin
      trouve:=4;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)],valaxeR1)<>0)
      and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<3)
    then
     begin
      trouve:=3;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)],valaxeR1)<>0)
      and (codeaxeR3='') and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<2)
    then
     begin
      trouve:=2;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (pos(TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)],valaxeR1)<>0)
      and (codeaxeR2='') and (codeaxeR3='') and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<1)
    then
     begin
      trouve:=1;
      trouveLoi:=Loi;
     end;

{   //on tombe sur une loi particuliere
   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (valaxeR6=TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)])
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (valaxeR7=TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)])
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (valaxeR8=TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)])
      and (codeaxeR9<>'') and (TStCodeAxe.IndexOf(codeaxeR9)<>-1)
      and (valaxeR9=TabValAxe[TStCodeAxe.IndexOf(codeaxeR9)])
      and (codeaxeR10<>'') and (TStCodeAxe.IndexOf(codeaxeR10)<>-1)
      and (valaxeR10=TabValAxe[TStCodeAxe.IndexOf(codeaxeR10)])
      and (trouve<10)
    then
     begin
      trouve:=10;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (valaxeR6=TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)])
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (valaxeR7=TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)])
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (valaxeR8=TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)])
      and (codeaxeR9<>'') and (TStCodeAxe.IndexOf(codeaxeR9)<>-1)
      and (valaxeR9=TabValAxe[TStCodeAxe.IndexOf(codeaxeR9)])
      and (codeaxeR10='')
      and (trouve<9)
    then
     begin
      trouve:=9;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (valaxeR6=TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)])
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (valaxeR7=TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)])
      and (codeaxeR8<>'') and (TStCodeAxe.IndexOf(codeaxeR8)<>-1)
      and (valaxeR8=TabValAxe[TStCodeAxe.IndexOf(codeaxeR8)])
      and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<8)
    then
     begin
      trouve:=8;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (valaxeR6=TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)])
      and (codeaxeR7<>'') and (TStCodeAxe.IndexOf(codeaxeR7)<>-1)
      and (valaxeR7=TabValAxe[TStCodeAxe.IndexOf(codeaxeR7)])
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<7)
    then
     begin
      trouve:=7;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6<>'') and (TStCodeAxe.IndexOf(codeaxeR6)<>-1)
      and (valaxeR6=TabValAxe[TStCodeAxe.IndexOf(codeaxeR6)])
      and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<6)
    then
     begin
      trouve:=6;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5<>'') and (TStCodeAxe.IndexOf(codeaxeR5)<>-1)
      and (valaxeR5=TabValAxe[TStCodeAxe.IndexOf(codeaxeR5)])
      and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<5)
    then
     begin
      trouve:=5; 
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4<>'') and (TStCodeAxe.IndexOf(codeaxeR4)<>-1)
      and (valaxeR4=TabValAxe[TStCodeAxe.IndexOf(codeaxeR4)])
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<4)
    then 
     begin
      trouve:=4;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3<>'') and (TStCodeAxe.IndexOf(codeaxeR3)<>-1)
      and (valaxeR3=TabValAxe[TStCodeAxe.IndexOf(codeaxeR3)])
      and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<3)
    then
     begin
      trouve:=3;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2<>'') and (TStCodeAxe.IndexOf(codeaxeR2)<>-1)
      and (valaxeR2=TabValAxe[TStCodeAxe.IndexOf(codeaxeR2)])
      and (codeaxeR3='') and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<2)
    then 
     begin
      trouve:=2;
      trouveLoi:=Loi;
     end;

   if (codeaxeR1<>'') and (TStCodeAxe.IndexOf(codeaxeR1)<>-1)
      and (valaxeR1=TabValAxe[TStCodeAxe.IndexOf(codeaxeR1)])
      and (codeaxeR2='') and (codeaxeR3='') and (codeaxeR4='')
      and (codeaxeR5='') and (codeaxeR6='') and (codeaxeR7='')
      and (codeaxeR8='') and (codeaxeR9='') and (codeaxeR10='')
      and (trouve<1)
    then 
     begin
      trouve:=1;
      trouveLoi:=Loi;
     end;

}

   //cas on n'a pas encore trouvé de loi
   //et on tombe sur la loi générale
   if (codeAxeR1='') and (trouve=-1)
    then 
     begin
      trouve:=0;
      trouveLoi:=Loi;
     end;

   Q.next;
  end;
 ferme(Q);

 //si on a trouvé une loi
 //on recupère les %
 if trouveLoi<>''
  then
   begin
    if OkTaille
     then RemplitTabPrctEclatTaille(trouveLoi,tabValeurPrct,DateDebLoi,LoiGrilleTaille)
     else RemplitTabPrctEclat(codeSession,trouveLoi,maille,mode,tabValeurPrct,DateDebLoi,DateFinLoi);
    result:=true;
   end;

end;

//calcul d'éclatement par delai ou par taille
procedure RemplitArbreNiveauEclatement(const codeSession,TypeEclat:hString);
var sreq,codesql,codesqlH,noeud:hString;
    nivMax,nivS,i,k,NewNoeud:integer;
    Q,Q2:TQuery;
    TabValAxe:array [1..12] of hString;
    newval,newQteVal,newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAval6:double;
    histoval,Qteval,CAval2,CAval3,CAval4,CAval5,CAval6,Prevu,Histo,Realise,QteRet:double;
    NomValeurAxe:hString;
    ValeurDelai:TDateTime;
    maille,LoiGrilleTaille,TitreNiveau:hString;
    tabValeurPrct:array [0..52] of double;
    DateDebLoi,DateFinLoi:TDateTime;
    TStLCodeAxe:TStringList;
    OkLoi,OkTaille:boolean;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    TabTaille:array [0..19] of hString;
    codeart,GrilleTaille:hString;
    ListeQte,ListeQteRes: array[0..19] of double;
    rapport,QteR,QteRArrondi,errQteR:double;
    QteValArrondi,QteValX,errQteVal:double;
    QteP,QtePArrondi,errQteP:double;
    Enr:TEnreg;
    CoeffRet:double;
begin
 OkTaille:=false;
 TitreNiveau:='DELAI';
 if TypeEclat='TAILLE'
  then OkTaille:=true;

 //cherche le nombre de niveau definit
 nivS:=1;
 if OkSessionObjectif(codeSession)
  then nivMax:=ChercheNivMax(codeSession)
  else
   begin
    //niveau prevision
    nivMax:=ChercheNivMaxSession(codeSession);
    if (TypeEclat='DELAI') and (SessionDelai(codeSession))
     then nivMax:=nivMax-1;
    //niveau session
    nivS:=ChercheNivMax(codeSession);
   end;

 //initialise code sql
 codesql:='QBR_NUMNOEUD,QBR_VALEURAXE,'+
          'QBR_OP1,QBR_QTEC,QBR_OP2,QBR_OP3,'+
          'QBR_OP4,QBR_OP5,QBR_OP6,'+
          'QBR_REF1,QBR_QTEREF,QBR_REF2,QBR_REF3,'+
          'QBR_REF4,QBR_REF5,QBR_REF6,QBR_PREVU,'+
          'QBR_HISTO,QBR_REALISE,QBR_QTERETENUE ';

 for i:=1 to NivMax-1 do
  begin
   codesql:=codesql+',QBR_VALAXENIV'+IntToStr(i);
  end;

 codesql:=codesql+',QBR_COEFFRETENU ';

 //suppression de l'ancien eclat
 //dans le cas delai car on rajoute un niveau à l'arbre
 //pas dans le cas taille car on ne rajoute pas de niveau
 //on met à jour les 20 qtés du dernier niveau
 if TypeEclat='DELAI'
 then
 begin
 MExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NIVEAU="'+IntToStr(NivMax+1)+'"',
                   'BPEclatement (RemplitArbreNiveauEclatement).');
 MExecuteSql('DELETE FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"',
                   'BPEclatement (RemplitArbreNiveauEclatement).');
 end;

 //Nouveau noeud à rajouter
 NewNoeud:=BPIncrementenumNoeud(codeSession);


 //code axe session
 TStLCodeAxe:=TStringList.Create;
 ChercheCodeAxeSessionTStL(codeSession,TStLCodeAxe);

 //cherche période session
 ChercheDateDDateFPeriode(codeSession,
                          DateDebC,DateFinC,DateDebRef,DateFinRef);

 //sql
 Q:=MOPenSql('SELECT '+codesql+' FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NIVEAU="'+IntToStr(NivMax)+'"',
             'BPEclatement (RemplitArbreNiveauEclatement).',true);

 InitMoveProgressForm(nil, TraduireMemoire('Session : ')+codeSession , TraduireMemoire('Veuillez patienter'),Q.RecordCount , True, True);

 while not Q.eof do
  begin
   if not MoveCurProgressForm(TraduireMemoire('Calcul en cours...')) then
   begin
    MoveCurProgressForm(TraduireMemoire('Annulation en cours...'));
    break;
   end;

   MoveCurProgressForm(TraduireMemoire('Calcul en cours...'));
   noeud:=Q.fields[0].asString;
   newval:=Q.fields[2].asFloat;
   newQteval:=Q.fields[3].asFloat;
   newCAval2:=Q.fields[4].asFloat;
   newCAval3:=Q.fields[5].asFloat;
   newCAval4:=Q.fields[6].asFloat;
   newCAval5:=Q.fields[7].asFloat;
   newCAval6:=Q.fields[8].asFloat;
   histoval:=Q.fields[9].asFloat;
   Qteval:=Q.fields[10].asFloat;
   CAval2:=Q.fields[11].asFloat;
   CAval3:=Q.fields[12].asFloat;
   CAval4:=Q.fields[13].asFloat;
   CAval5:=Q.fields[14].asFloat;
   CAval6:=Q.fields[15].asFloat;
   Prevu:=Q.fields[16].asFloat;
   Histo:=Q.fields[17].asFloat;
   Realise:=Q.fields[18].asFloat;
   QteRet:=Q.fields[19].asFloat;

   for i:=1 to NivMax-1 do
    TabValAxe[i]:=Q.fields[19+i].asString;

   CoeffRet:=Q.fields[19+NivMax].asFloat;

   TabValAxe[NivMax]:=Q.fields[1].asString;

   if TypeEclat='TAILLE'
    then
     begin
      //valeur du code article
      //dans l'arbre le code article se trouve obligatoirement
      //sur le niveau suivant le dernier niveau de la session
      codeArt:=TabValAxe[nivS+1];
      //grille de taille de l'article
      if dm_trouveEnr(idm_tmpBPArticleTai,[codeArt],Enr)=0
       then GrilleTaille:=Enr.ch(BPARTICLETAI_GRILLETAI);
     end;

   for i:=0 to 52 do tabValeurPrct[i]:=0;

   //cherche tab prct loi
   OkLoi:=RemplitTabPrctEclatLoi(codeSession,GrilleTaille,OkTaille,TabValAxe,TStLCodeAxe,
                          maille,LoiGrilleTaille,tabValeurPrct,DateDebloi,DateFinLoi);
   errQteR:=0;
   errQteVal:=0;
   errQteP:=0;
   //si on a trouvé une loi qui correspond
   if OkLoi
    then
     begin
      //cas delai
      //on insert un nouvel enreg sous le dernier niveau de la session
      if TypeEclat='DELAI'
       then
        begin
         //cherche le nombre d'intervalle entre 2 dates pour une maille donnée
         k:=DonneNbIntervalleMailleDateDebDateFin(maille,DateDebC,DateFinC);
         //pour chaque intervalle
         //on insert un nouvel enreg
         for i:=0 to k-1 do
          begin
           NewNoeud:=NewNoeud+1;
           //cherche le nom correspondant à l'intervalle suivant la maille
           NomValeurAxe:=DateTimeToStr(ChercheNomValeurAxeDelai(maille,i,DateDebC));
           ValeurDelai:=ChercheNomValeurAxeDelai(maille,i,DateDebC);

           QteR:=QteRet*tabValeurPrct[i];
           QteRArrondi:=round(QteR+errQteR);
           errQteR:=errQteR+(QteR-QteRArrondi);

           QteP:=prevu*tabValeurPrct[i];
           QtePArrondi:=round(QteP+errQteP);
           errQteP:=errQteP+(QteP-QtePArrondi);

           QteValX:=newQteVal*tabValeurPrct[i];
           QteValArrondi:=round(QteValX+errQteVal);
           errQteVal:=errQteVal+(QteValX-QteValArrondi);

           MExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,'+
                    'QBR_NUMNOEUDPERE,QBR_VALEURAXE,QBR_NIVEAU,'+
                    'QBR_CODEAXE,QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,'+
                    'QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
                    'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,'+
                    'QBR_DATEDELAI,'+
                    'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                    'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                    'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                    'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                    'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                    'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                    'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                    'QBR_HISTO,QBR_REALISE,QBR_PREVU,QBR_QTERETENUE,QBR_MAILLE,'+
                    'QBR_VALBLOQUE,QBR_VALBLOQUETMP,QBR_VALBLOQUETMP1,QBR_COEFFRETENU'+
                    ') VALUES ("'+codeSession+
                    '",'+intToStr(NewNoeud)+','+noeud+',"'+NomValeurAxe+
                    '",'+intToStr(nivMax+1)+',"'+TitreNiveau+'","'+TabValAxe[1]+
                    '","'+TabValAxe[2]+'","'+TabValAxe[3]+'","'+TabValAxe[4]+
                    '","'+TabValAxe[5]+'","'+TabValAxe[6]+'","'+TabValAxe[7]+
                    '","'+TabValAxe[8]+'","'+TabValAxe[9]+'","'+
                    USDATETIME(ValeurDelai)+'",'+
                    STRFPOINT(newval*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(histoval*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(newCAVal2*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(CAVal2*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(newCAVal3*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(CAVal3*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(newCAVal4*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(CAVal4*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(newCAVal5*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(CAVal5*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(newCAVal6*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(CAVal6*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(QteValArrondi)+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(QteVal*tabValeurPrct[i])+','+STRFPOINT(tabValeurPrct[i]*100)+','+
                    STRFPOINT(Histo*tabValeurPrct[i])+','+STRFPOINT(Realise*tabValeurPrct[i])+','+
                    STRFPOINT(QtePArrondi)+','+STRFPOINT(QteRArrondi)+',"'+
                    maille+'","-","-","-",'+STRFPOINT(CoeffRet)+')',
                    'BPEclatement (RemplitArbreNiveauEclatement).');
          end
        end
       else
        begin
         //cas taille

         //remplit une liste avec les qtés par taille
         //gestion des arrondis et prorata)
         rapport:=0;
         if Prevu<>0
          then rapport:=QteRet/100;
         for i:=0 to 19 do
          ListeQte[i]:= tabValeurPrct[i]*100;

         BPCalculQteListProrata(0,ListeQte,Rapport,ListeQteRes);
         sReq:='';
         for i:= 1 to 19 do
           sReq:=sReq+',QBR_QTET'+intToStr(i+1)+'='+strFPoint(ListeQteRes[i]);
         //mise à jour de l'enreg
         MExecuteSql('UPDATE QBPARBRE SET QBR_QTET1='+strFPoint(ListeQteRes[0])+
                      sReq+
                     ' WHERE QBR_CODESESSION="'+codeSession+
                     '" AND QBR_NUMNOEUD="'+noeud+'" ',
                    'BPEclatement (RemplitArbreNiveauEclatement).');

        end;
     end;
   Q.next;
  end;
 ferme(Q);


 if not MoveCurProgressForm(TraduireMemoire('Calcul en cours...')) then
 begin
   ExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"');
   ExecuteSql('DELETE FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"');
   ExecuteSQL('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="-",QBS_BPINITIALISE="" WHERE QBS_CODESESSION="'+codeSession+'"');
   PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), 'Session : '+codeSession);
 end;

 FiniMoveProgressForm;

 TStLCodeAxe.free;
end;


//lance le calcul d'éclatement
//par delai ou par taille
//typeEclat=DELAI ou TAILLE
procedure CalculEclatementLoi(const codeSession,TypeEclat:hString);
var codeSql:hString;
begin
 RemplitArbreNiveauEclatement(codeSession,TypeEclat);

 //mise à jour dans qbpsessionbp session éclatée par delai ou par taille
 if TypeEclat='DELAI'
  then codeSql:=' QBS_SESSIONECLAT="X" '
  else codeSql:=' QBS_SESSIONECLATAI="X" ';
 MExecuteSql('UPDATE QBPSESSIONBP SET '+codeSql+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPEclatement (CalculEclatementLoi).');
 ModifSession(codeSession,'X');
end;


//************************************************************************
//
//                       Dans les lois d'eclatement
//                       Intialisation des tableaux avec les %
//                       par rapport à un histo
//
//************************************************************************

//intialisation des % par rapport à une période de reference de la loi delai
//cas Orli
procedure InitialisationPrctDelaiParRapportHistoO(const mode,codeRestriction:hString;
          dateDeb,dateFin:TDateTime;
          var TabQteLoi:array of double);
var Q:TQuery;
    codeSql,codeSqlQ1,codeSqlQ2:hString;
    codeI,indice,AnneeI,AN,SemDep:integer;
    Qte,QteTotale:double;
    anDep,moisDep,jourDep:word;
    numT,anT,numDR,anDR:integer;
    Per:TPeriodeP;
    DateI:TDateTime;
begin
 codeSql:='';
 codeSqlQ1:='';
 codeSqlQ2:='';

 //cherche le code sql suivant le mode d'initialisation    
 //par quadrimestre
 if mode='7'
  then codeSql:='QBV_DATEPIVOT ';
 //par trimestre
 if mode='6'
  then codeSql:='QBV_DATEPIVOT ';
 //par mois4-4-5
 if mode='5'
  then codeSql:='QBV_DATEPIVOT ';
 //par mois
 if mode='4'
  then codeSql:='MONTH(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';
 //par quinzaine
 if mode='3'
  then
   begin
    codeSql:='MONTH(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';
    codeSqlQ1:=' AND DAY(QBV_DATEPIVOT)<15 ';
    codeSqlQ2:=' AND DAY(QBV_DATEPIVOT)>=15 ';
   end;
 //par semmaine
 if mode='2'
  then codeSql:='WEEK(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';

 //indice de départ
 DecodeDate(dateDeb,anDep,moisDep,jourDep);
 SemDep:= numSemaine(dateDeb,An);

 QteTotale:=1;
 //récupère qté totale pour calcul du %
 Q:=MOpenSql('SELECT SUM(QBV_QTE1) FROM QBPPIVOT '+
             'WHERE QBV_DATEPIVOT>="'+USDATETIME(DateDeb)+
             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFin)+
             '" '+codeRestriction,
             'BPEclatement (InitialisationPrctDelaiParRapportHistoO).',true);
 if not Q.eof
  then QteTotale:=valeurR(Q.fields[0].asString);
 ferme(Q);

 CodeI:=0;
 AnneeI:=0;
 DateI:=0;

 Q:=MOpenSql('SELECT SUM(QBV_QTE1),'+codeSql+' FROM QBPPIVOT '+
             'WHERE QBV_DATEPIVOT>="'+USDATETIME(DateDeb)+
             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFin)+
             '" '+codeSqlQ1+codeRestriction+
             'GROUP BY '+codeSql,
             'BPEclatement (InitialisationPrctDelaiParRapportHistoO).',true);
 while not Q.eof do
  begin
   //récupère les données
   Qte:=valeurR(Q.fields[0].asString);
   if (mode='5') or (mode='6') or (mode='7')
    then
     begin
      DateI:=Q.fields[1].asDateTime;
     end
    else
     begin
      CodeI:=valeurI(Q.fields[1].asString);
      AnneeI:=valeurI(Q.fields[2].asString);
     end;

   //cherche indice tableau
   indice:=1;
   //par quadrimestre
   if mode='7'
    then
     begin
      DonneNumQuadrimestre(DateI,numT,anT);
      DonneNumQuadrimestre(dateDeb,numDR,anDR);
      if (anDR=anT)
       then indice:=numT-numDR+1
       else indice:=numT+3-numDR+1;
     end;
   //par trimestre
   if mode='6'
    then
     begin
      DonneNumTrimestre(DateI,numT,anT);
      DonneNumTrimestre(dateDeb,numDR,anDR);
      if (anDR=anT)
       then indice:=numT-numDR+1
       else indice:=numT+4-numDR+1;
     end;
   //par mois 4-4-5
   if mode='5'
    then
     begin
      Per:=FindPeriodeP(DateI);
      indice:=Per.Periode;
     end;
   //par mois
   if mode='4'
    then
     begin
      if (CodeI>MoisDep) and (anneeI=anDep)
       then indice:=CodeI-MoisDep+1;
      if (CodeI<MoisDep) and (anneeI>anDep)
       then indice:=CodeI+(12-MoisDep)+1;
     end;
   //par quinzaine
   if mode='3'
    then
     begin
      indice:=CodeI;
     end;
   //par semaine
   if mode='2'
    then
     begin
      if (CodeI>SemDep) and (anneeI=anDep)
       then indice:=CodeI-SemDep-1;
      if (CodeI<SemDep) and (anneeI>anDep)
       then indice:=CodeI+(52-SemDep)+1;
     end;

   //insertion dans le tableau de la loi
   //par qudrimestre
   if mode='7'
    then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;   
   //par trimestre
   if mode='6'
    then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;
   //par mois 4-4-5
   if mode='5'
    then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;

   //par mois
   if mode='4'
    then TabQteLoi[indice]:=(Qte/QteTotale)*100;
   //par quinzaine
   if mode='3'
    then TabQteLoi[CodeI+(CodeI-1)]:=(Qte/QteTotale)*100;
   //par semaine
   if mode='2'
    then TabQteLoi[indice]:=(Qte/QteTotale)*100;
   Q.next;
  end;
 ferme(Q);

 if mode='3'
  then
   begin
    //dans le cas mode =3 quinzaine il faut chercher les 2° quinzaine dans le mois
    Q:=MOpenSql('SELECT SUM(QBV_QTE1),'+codeSql+' FROM QBPPIVOT '+
                'WHERE QBV_DATEPIVOT>="'+USDATETIME(DateDeb)+
                '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFin)+
                '" '+codeSqlQ2+codeRestriction+
                'GROUP BY '+codeSql,
                'BPEclatement (InitialisationPrctDelaiParRapportHistoO).',true);
    while not Q.eof do
     begin
      //récupère les données
      Qte:=valeurR(Q.fields[0].asString);
      CodeI:=valeurI(Q.fields[1].asString);

      //insertion dans le tableau de la loi
      if mode='3'
       then TabQteLoi[codeI*2]:=(Qte/QteTotale)*100;
      Q.next;
     end;
    ferme(Q);
   end;
end;


//intialisation des % par rapport à une période de reference de la loi delai
//cas PGI
procedure InitialisationPrctDelaiParRapportHistoPgi(const mode,codeSession,codeRestriction,codeJoinR:hString;
          dateDebRef,dateFinRef:TDateTime;
          var TabQteLoi:array of double);
var Q:TQuery;
    codeSql,codeSqlQ1,codeSqlQ2,reqSql:hString;
    codeI,indice,AnneeI,AN,SemDep:integer;
    Qte,QteTotale:double;
    anDep,moisDep,jourDep:word;
    DateI:TDateTime;
    per: TperiodeP;
    numT,anT,numDR,anDR:integer;
    NbValAff :integer;
    codeChpsSum,Prefixe,codeTable,_tmp_ : String;
begin
  codeSql:='';
  codeSqlQ1:='';
  codeSqlQ2:='';

  //cherche le code sql suivant le mode d'initialisation
  //par quadrimestre
  if mode='7' then codeSql:='[DATEDEBUT] ';
  //par trimestre
  if mode='6' then codeSql:='[DATEDEBUT] ';
  //par mois 4-4-5
  if mode='5' then codeSql:='[DATEDEBUT] ';
  //par mois
  if mode='4' then codeSql:='MONTH([DATEDEBUT]),YEAR([DATEDEBUT]) ';
  //par quinzaine
  if mode='3' then
  begin
    codeSql:='MONTH([DATEDEBUT]),YEAR([DATEDEBUT]) ';
    codeSqlQ1:=' AND DAY([DATEDEBUT])<15 ';
    codeSqlQ2:=' AND DAY([DATEDEBUT])>15 ';
  end;
  //par semmaine
  if mode='2' then codeSql:='WEEK([DATEDEBUT]),YEAR([DATEDEBUT]) ';

  //indice de départ
  DecodeDate(dateDebRef,anDep,moisDep,jourDep);
  SemDep:=numSemaine(dateDebRef,An);

  QteTotale:=1;
   //récupère qté totale pour calcul du %
  ReqSql := 'SELECT [CHAMP] FROM [TABLE]'+codejoinR+
            ' WHERE [DATEDEBUT]>="'+USDATETIME(DateDebRef)+
            '" AND [DATEFIN]<="'+USDATETIME(DateFinRef)+
            '" '+codeRestriction;
  Case ContextBP of
    0,1 : begin
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','GL');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','GL_DATEPIECE');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','GL_DATEPIECE');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','LIGNE');
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(GL_QTEFACT)');
        end;
    2 : begin
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','Y');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','Y_DATECOMPTABLE');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','Y_DATECOMPTABLE');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','ANALYTIQ');
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(Y_DEBIT)');
        end;
    3 : begin
          DonnePrefixe(codeSession,Prefixe,_tmp_);
          if ((Prefixe = 'NBS') OR (Prefixe = 'EFM')) then Prefixe := 'PPU';

          ReqValAff(codeSession,'',0,NbValAff,codeChpsSum,_tmp_,codeTable,_tmp_,_tmp_,_tmp_,_tmp_,_tmp_);
          Delete(codeChpsSum, Pos(',',codeChpsSum),length(codeChpsSum));
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]',Prefixe);
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]',Prefixe+'_DATEDEBUT');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]',Prefixe+'_DATEFIN');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]',codeTable);
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]',codeChpsSum);
        end;
  end; //CASE
  Q:=MOpenSql(ReqSql,'BPEclatement (InitialisationPrctDelaiParRapportHistoPgi).',true);
  if not Q.eof then QteTotale:=valeurR(Q.fields[0].asString);
  ferme(Q);

  CodeI:=0;
  AnneeI:=0;
  DateI:=0;

  ReqSql := 'SELECT [CHAMP],'+codeSql+' FROM [TABLE]'+codejoinR+
            ' WHERE [DATEDEBUT]>="'+USDATETIME(DateDebRef)+
            '" AND [DATEFIN]<="'+USDATETIME(DateFinRef)+
            '" '+codeSqlQ1+codeRestriction+
            'GROUP BY '+codeSql;
  Case ContextBP of
    0,1 : begin
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','GL');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','GL_DATEPIECE');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','GL_DATEPIECE');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','LIGNE');
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(GL_QTEFACT)');
        end;
    2 : begin
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','Y');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','Y_DATECOMPTABLE');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','Y_DATECOMPTABLE');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','ANALYTIQ');
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(Y_DEBIT)');
        end;
    3 : begin
          ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]',Prefixe);
          ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]',Prefixe+'_DATEDEBUT');
          ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]',Prefixe+'_DATEFIN');
          ReqSql := AnsiReplaceText(ReqSql,'[TABLE]',codeTable);
          ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]',codeChpsSum);
        end;
  end; //CASE
  Q:=MOpenSql(ReqSql,'BPEclatement (InitialisationPrctDelaiParRapportHistoPgi).',true);
  while not Q.eof do
  begin
    //récupère les données
    Qte:=valeurR(Q.fields[0].asString);
    if (mode='5') or (mode='6') or (mode='7') then DateI:=Q.fields[1].asDateTime
    else
    begin
      CodeI:=valeurI(Q.fields[1].asString);
      AnneeI:=valeurI(Q.fields[2].asString);
    end;

    //cherche indice tableau
    indice:=1;

    //par quadrimestre
    if mode='7' then
    begin
      DonneNumQuadrimestre(DateI,numT,anT);
      DonneNumQuadrimestre(dateDebRef,numDR,anDR);
      if (anDR=anT) then indice:=numT-numDR+1
      else indice:=numT+3-numDR+1;
    end;
    //par trimestre
    if mode='6' then
    begin
      DonneNumTrimestre(DateI,numT,anT);
      DonneNumTrimestre(dateDebRef,numDR,anDR);
      if (anDR=anT) then indice:=numT-numDR+1
      else indice:=numT+4-numDR+1;
    end;
    //par mois 4-4-5
    if mode='5' then
    begin
      Per:=FindPeriodeP(DateI);
      indice:=Per.Periode;
    end;
    //par mois
    if mode='4' then
    begin
      if (CodeI>MoisDep) and (anneeI=anDep)
      then indice:=CodeI-MoisDep+1;
      if (CodeI<MoisDep) and (anneeI>anDep)
      then indice:=CodeI+(12-MoisDep)+1;
    end;
    //par quinzaine
    if mode='3' then
    begin
      if moisDep=1 then indice:=CodeI+(CodeI-1)
      else
      begin
        if (codeI>MoisDep) and (anneeI=anDep)
        then indice:=(CodeI-MoisDep+1)+(CodeI-MoisDep);
        if (codeI<MoisDep) and (anneeI>anDep)
        then indice:=12+CodeI+(CodeI-1);
      end;
    end;
    //par semaine
    if mode='2' then
    begin
      if SemDep=1 then indice:=CodeI
      else
      begin
        if (CodeI>SemDep) and (anneeI=anDep)
        then indice:=CodeI-SemDep;
        if (CodeI<SemDep) and (anneeI>anDep)
        then indice:=CodeI+(52-SemDep)+1;
      end;
    end;

    //insertion dans le tableau de la loi
    //par qudrimestre
    if mode='7' then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;
    //par trimestre
    if mode='6' then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;
    //par mois 4-4-5
    if mode='5' then TabQteLoi[indice]:=TabQteLoi[indice]+(Qte/QteTotale)*100;
    //par mois
    if mode='4' then TabQteLoi[indice]:=(Qte/QteTotale)*100;
    //par quinzaine
    if mode='3' then TabQteLoi[indice]:=(Qte/QteTotale)*100;
    //par semaine
    if mode='2' then TabQteLoi[indice]:=(Qte/QteTotale)*100;
    Q.next;
  end;
  ferme(Q);

  if mode='3' then
  begin
    //dans le cas mode =3 quinzaine il faut chercher les 2° quinzaine dans le mois
    ReqSql := 'SELECT [CHAMP],'+codeSql+' FROM [TABLE]'+codeJoinR+
              ' WHERE [DATEDEBUT]>="'+USDATETIME(DateDebRef)+
              '" AND [DATEFIN]<="'+USDATETIME(DateFinRef)+
              '" '+codeSqlQ2+codeRestriction+
              'GROUP BY '+codeSql;
    Case ContextBP of
      0,1 : begin
            ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','GL');
            ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','GL_DATEPIECE');
            ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','GL_DATEPIECE');
            ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','LIGNE');
            ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(GL_QTEFACT)');
          end;
      2 : begin
            ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]','Y');
            ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]','Y_DATECOMPTABLE');
            ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]','Y_DATECOMPTABLE');
            ReqSql := AnsiReplaceText(ReqSql,'[TABLE]','ANALYTIQ');
            ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]','SUM(Y_DEBIT)');
          end;
      3 : begin
            ReqSql := AnsiReplaceText(ReqSql,'[PREFIXE]',Prefixe);
            ReqSql := AnsiReplaceText(ReqSql,'[DATEDEBUT]',Prefixe+'_DATEDEBUT');
            ReqSql := AnsiReplaceText(ReqSql,'[DATEFIN]',Prefixe+'_DATEFIN');
            ReqSql := AnsiReplaceText(ReqSql,'[TABLE]',codeTable);
            ReqSql := AnsiReplaceText(ReqSql,'[CHAMP]',codeChpsSum);
          end;
    end; //CASE
    Q:=MOpenSql(ReqSql,'BPEclatement (InitialisationPrctDelaiParRapportHistoPgi).',true);
    while not Q.eof do
    begin
      //récupère les données
      Qte:=valeurR(Q.fields[0].asString);
      CodeI:=valeurI(Q.fields[1].asString);
      AnneeI:=valeurI(Q.fields[2].asString);

      //insertion dans le tableau de la loi
      if moisDep=1 then indice:=codeI*2
      else
      begin
        indice:=1;
        if (codeI>=MoisDep) and (anneeI=anDep)
        then indice:=(CodeI-MoisDep+1)*2;
        if (codeI<MoisDep) and (anneeI>anDep)
        then indice:=(CodeI+MoisDep-1)*2;
      end;
      TabQteLoi[indice]:=(Qte/QteTotale)*100;
      Q.next;
    end;
    ferme(Q);
  end;
end;

//intialisation des % par rapport à une période de reference dans la loi delai
procedure InitialisationPrctDelaiParRapportHisto(const mode,codeSession,codeRestriction,codeJoinR:hString;
          dateDebRef,dateFinRef:TDateTime;
          var TabQteLoi:array of double);
begin
 if not BPOkOrli
  then InitialisationPrctDelaiParRapportHistoPgi(mode,codeSession,codeRestriction,codeJoinR,
        dateDebRef,dateFinRef,TabQteLoi)
  else //-----------------> ORLI
       InitialisationPrctDelaiParRapportHistoO(mode,codeRestriction,
        dateDebRef,dateFinRef,TabQteLoi);
       //ORLI <-----------------
end;


//initialisation des % par rappport à une periode de reference dans la loi taille
procedure InitialisationPrctTailleParRapportHisto(const CodeTaille,codeRestriction:hString;
          dateDeb,dateFin:TDateTime;
          var TabQteLoi:array of double);
var Q:TQuery;
    QteTotale,somme:double;
    i:integer;
begin
 QteTotale:=1;
 //récupère qté totale pour calcul du %
 Q:=MOpenSql('SELECT SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
             'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+'+
             'QBV_QTET10+QBV_QTET11+QBV_QTET12+QBV_QTET13+'+
             'QBV_QTET14+QBV_QTET15+QBV_QTET16+QBV_QTET17+'+
             'QBV_QTET18+QBV_QTET19+QBV_QTET20) FROM QBPPIVOT '+
             'WHERE QBV_DATEPIVOT>="'+USDATETIME(DateDeb)+
             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFin)+
             '" AND QBV_VALAXEP5="'+CodeTaille+'" '+codeRestriction,
             'BPEclatement (InitialisationPrctTailleParRapportHisto).',true);
 if not Q.eof
  then QteTotale:=valeurR(Q.fields[0].asString);
 ferme(Q);
 for i:=1 to 20 do
  TabQteLoi[I]:=0;
 if qteTotale<>0
  then
   begin
    //remplit le tableau des 20 qtés de 20 tailles de la grille de taille
    Q:=MOpenSql('SELECT SUM(QBV_QTET1),SUM(QBV_QTET2),SUM(QBV_QTET3),'+
                'SUM(QBV_QTET4),SUM(QBV_QTET5),SUM(QBV_QTET6),'+
                'SUM(QBV_QTET7),SUM(QBV_QTET8),SUM(QBV_QTET9),'+
                'SUM(QBV_QTET10),SUM(QBV_QTET11),SUM(QBV_QTET12),'+
                'SUM(QBV_QTET13),SUM(QBV_QTET14),SUM(QBV_QTET15),'+
                'SUM(QBV_QTET16),SUM(QBV_QTET17),SUM(QBV_QTET18),'+
                'SUM(QBV_QTET19),SUM(QBV_QTET20) FROM QBPPIVOT '+
                'WHERE QBV_DATEPIVOT>="'+USDATETIME(DateDeb)+
                '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFin)+
                '" AND QBV_VALAXEP5="'+CodeTaille+'" '+codeRestriction,
                'BPEclatement (InitialisationPrctTailleParRapportHisto).',true);
    if not Q.eof
     then
      begin
       somme:=0;
       for i:=1 to 20 do
        begin
         TabQteLoi[I]:=(valeurR(Q.fields[I-1].asString)/QteTotale)*100;
         somme:=somme+VALEUR(Format_('%10.4f',[tabQteLoi[i]]));   
        end;
       TabQteLoi[1]:=TabQteLoi[1]+(100-somme);
      end;
    ferme(Q);
  end;
end;

end.

