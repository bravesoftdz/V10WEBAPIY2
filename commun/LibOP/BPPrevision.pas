unit BPPrevision;

interface

uses HCtrls,SysUtils;

procedure RemplitQBPBudgetprevAvecCoeff(const codeSession:hString);
procedure ModifQBPBudgetprevCoeff(const codeSession:hString;
          coeff:double;
          TabValeurAxe:array of hString);

implementation

uses uutil,
     BPFctSession;
{
procedure RemplitQBPBudgetprevAvecCoeff(const codeSession:hString);
var DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin:TDateTime;
    SaiCmd,SaiCmdRef:hString;
    code,iSt:hString;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    nivMax,j:integer;
    codeRestriction,codeWhere,codeValAxe8,codeJoinR:hString;
begin
 ExecuteSql('DELETE FROM QBPBUDGETPREV '+
              'WHERE QBB_CODESESSION="'+codeSession+'"');

 //recupère les données concernant la session
 ChercheDateDDateFPeriode(codeSession,DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin);
 ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 //cherche les codes sql
 //on prend les codes jusqu'à nivmax définit
 //pour qbpbudgetprev
 code:='';
 for j:=1 to NivMax+1 do
  begin
   ist:=IntToStr(j);
   if TabNumValAxe[j]<>''
    then code:=code+' AND QBV_VALAXEP'+TabNumValAxe[j]+'=QBT_CODEAXEC'+IntToStr(j-1);
  end;

 codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,codeJoinR);

 //insert dans qbpbudgetprev
 codeWhere:=' QBV_DATEPIVOT>="'+USDATETIME(DatePerCDeb)+
            '" AND QBV_DATEPIVOT<="'+USDATETIME(DatePerCFin)+'" ';
 codeValAxe8:='QBV_VALAXEP8';
 if SaiCmd<>''
  then
   begin
    codeWhere:=' QBV_VALAXEP8="'+SaiCmd+'" ';
    codeValAxe8:='"'+SaiCmd+'"';
   end;

 MExecuteSql('INSERT INTO QBPBUDGETPREV '+
            '(QBB_CODESESSION,QBB_NUMORDRE,QBB_CODEIDENT,QBB_MODEGESTPREV,'+
            'QBB_NATURECMD,QBB_CODEMAILLE,QBB_DATEBP,QBB_EXTRAPOLABLE,'+
            'QBB_QTEBPT1,QBB_QTEBPT2,QBB_QTEBPT3,QBB_QTEBPT4,QBB_QTEBPT5,'+
            'QBB_QTEBPT6,QBB_QTEBPT7,QBB_QTEBPT8,QBB_QTEBPT9,QBB_QTEBPT10,'+
            'QBB_QTEBPT11,QBB_QTEBPT12,QBB_QTEBPT13,QBB_QTEBPT14,QBB_QTEBPT15,'+
            'QBB_QTEBPT16,QBB_QTEBPT17,QBB_QTEBPT18,QBB_QTEBPT19,QBB_QTEBPT20,'+
            'QBB_QTEBP1,QBB_QTEBP2,QBB_QTEBP3,QBB_QTEBP4,QBB_QTEBP5,QBB_QTEBP6,'+
            'QBB_VALAXEBP1,QBB_VALAXEBP2,QBB_VALAXEBP3,QBB_VALAXEBP4,'+
            'QBB_VALAXEBP5,QBB_VALAXEBP6,QBB_VALAXEBP7,QBB_VALAXEBP8,'+
            'QBB_VALAXEBP9,QBB_VALAXEBP10,QBB_VALAXEBP11,QBB_VALAXEBP12,'+
            'QBB_VALAXEBP13,QBB_VALAXEBP14,QBB_VALAXEBP15,QBB_VALAXEBP16,'+
            'QBB_VALAXEBP17,QBB_VALAXEBP18,QBB_VALAXEBP19,QBB_VALAXEBP20,'+
            'QBB_VALAXEBP21,QBB_VALAXEBP22,QBB_VALAXEBP23,QBB_VALAXEBP24,'+
            'QBB_VALAXEBP25,QBB_VALAXEBP26,QBB_VALAXEBP27,QBB_VALAXEBP28,'+
            'QBB_VALAXEBP29,QBB_VALAXEBP30,QBB_DEVISE,QBB_CODEPREVISION,'+
            'QBB_QTEBPC1,QBB_QTEBPC2,QBB_QTEBPC3,QBB_QTEBPC4,QBB_QTEBPC5,'+
            'QBB_QTEBPC6,QBB_QTEBPC7,QBB_QTEBPC8,QBB_QTEBPC9,QBB_QTEBPC10,'+
            'QBB_QTEBPC11,QBB_QTEBPC12,QBB_QTEBPC13,QBB_QTEBPC14,QBB_QTEBPC15,'+
            'QBB_QTEBPC16,QBB_QTEBPC17,QBB_QTEBPC18,QBB_QTEBPC19,QBB_QTEBPC20,QBB_CAC1,'+
            'QBB_DATEFIXATION) '+
            'SELECT "'+codeSession+'","'+'GLOBAL'+'",QBV_CODEIDENT,QBV_MODEGESTPREV,'+
            'QBV_NATURECMD,QBV_CODEMAILLE,QBV_DATEPIVOT,QBV_EXTRAPOLABLE,'+
            'IIF((QBV_QTET1*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET1*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET2*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET2*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET3*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET3*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET4*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET4*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET5*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET5*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET6*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET6*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET7*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET7*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET8*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET8*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET9*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET9*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET10*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET10*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET11*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET11*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET12*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET12*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET13*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET13*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET14*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET14*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET15*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET15*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET16*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET16*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET17*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET17*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET18*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET18*QBT_COEFFRETENU)+0.5,0)),'+
            'IIF((QBV_QTET19*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET19*QBT_COEFFRETENU)+0.5,0)),IIF((QBV_QTET20*QBT_COEFFRETENU<0.00001),0,ROUND((QBV_QTET20*QBT_COEFFRETENU)+0.5,0)),'+
            'QBV_QTE1*QBT_COEFFRETENU,QBV_QTE2,QBV_QTE3,QBV_QTE4,QBV_QTE5,QBV_QTE6,'+
            'QBV_VALAXEP1,QBV_VALAXEP2,QBV_VALAXEP3,QBV_VALAXEP4,'+
            'QBV_VALAXEP5,QBV_VALAXEP6,QBV_VALAXEP7,'+codeValAxe8+','+
            'QBV_VALAXEP9,QBV_VALAXEP10,QBV_VALAXEP11,QBV_VALAXEP12,'+
            'QBV_VALAXEP13,QBV_VALAXEP14,QBV_VALAXEP15,QBV_VALAXEP16,'+
            'QBV_VALAXEP17,QBV_VALAXEP18,QBV_VALAXEP19,QBV_VALAXEP20,'+
            'QBV_VALAXEP21,QBV_VALAXEP22,QBV_VALAXEP23,QBV_VALAXEP24,'+
            'QBV_VALAXEP25,QBV_VALAXEP26,QBV_VALAXEP27,QBV_VALAXEP28,'+
            'QBV_VALAXEP29,QBV_VALAXEP30,QBV_DEVISE,QBV_CODEPREVISION,'+
            'QBV_QTET1,QBV_QTET2,QBV_QTET3,QBV_QTET4,QBV_QTET5,'+
            'QBV_QTET6,QBV_QTET7,QBV_QTET8,QBV_QTET9,QBV_QTET10,'+
            'QBV_QTET11,QBV_QTET12,QBV_QTET13,QBV_QTET14,QBV_QTET15,'+
            'QBV_QTET16,QBV_QTET17,QBV_QTET18,QBV_QTET19,QBV_QTET20,QBV_QTE1,QBV_DATEFIXATION '+
            'FROM QBPPIVOT,QBPCOEFF WHERE '+codeWhere+
            ' AND QBT_CODESESSION="'+codeSession+'"'+code+codeRestriction,
            'BPrevision (RemplitQBPBudgetprevAvecCoeff).');
end;

}

procedure RemplitQBPBudgetprevAvecCoeff(const codeSession:hString);
var DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin:TDateTime;
    SaiCmd,SaiCmdRef:hString;
    code,iSt:hString;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    nivMax,j:integer;
    codeRestriction,codeWhere,codeValAxe8:hString;
begin
 ExecuteSql('DELETE FROM QBPBUDGETPREV '+
            'WHERE QBB_CODESESSION="'+codeSession+'"');

 //recupère les données concernant la session
 ChercheDateDDateFPeriode(codeSession,DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin);
 ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 //cherche les codes sql
 //on prend les codes jusqu'à nivmax définit
 //pour qbpbudgetprev
 code:=' AND QBR_NIVEAU="'+IntToStr(nivMax)+'" ';
 for j:=1 to NivMax do
  begin
   ist:=IntToStr(j);
   if TabNumValAxe[j]<>''
    then code:=code+' AND QBV_VALAXEP'+TabNumValAxe[j]+'=QBR_VALAXENIV'+IntToStr(j-1);
  end;

 if TabNumValAxe[NivMax+1]<>''
  then code:=code+' AND QBV_VALAXEP'+TabNumValAxe[NivMax]+'=QBR_VALEURAXE ';

 codeRestriction:='';//ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,codeJoinR);

 //insert dans qbpbudgetprev
 codeWhere:=' QBV_DATEPIVOT>="'+USDATETIME(DatePerCDeb)+
            '" AND QBV_DATEPIVOT<="'+USDATETIME(DatePerCFin)+'" ';
 codeValAxe8:='QBV_VALAXEP8';
 if SaiCmd<>''
  then
   begin
    codeWhere:=' QBV_VALAXEP8="'+SaiCmd+'" ';
    codeValAxe8:='"'+SaiCmd+'"';
   end;

 MExecuteSql('INSERT INTO QBPBUDGETPREV '+
            '(QBB_CODESESSION,QBB_NUMORDRE,QBB_CODEIDENT,QBB_MODEGESTPREV,'+
            'QBB_NATURECMD,QBB_CODEMAILLE,QBB_DATEBP,QBB_EXTRAPOLABLE,'+
            'QBB_QTEBPT1,QBB_QTEBPT2,QBB_QTEBPT3,QBB_QTEBPT4,QBB_QTEBPT5,'+
            'QBB_QTEBPT6,QBB_QTEBPT7,QBB_QTEBPT8,QBB_QTEBPT9,QBB_QTEBPT10,'+
            'QBB_QTEBPT11,QBB_QTEBPT12,QBB_QTEBPT13,QBB_QTEBPT14,QBB_QTEBPT15,'+
            'QBB_QTEBPT16,QBB_QTEBPT17,QBB_QTEBPT18,QBB_QTEBPT19,QBB_QTEBPT20,'+
            'QBB_QTEBP1,QBB_QTEBP2,QBB_QTEBP3,QBB_QTEBP4,QBB_QTEBP5,QBB_QTEBP6,'+
            'QBB_VALAXEBP1,QBB_VALAXEBP2,QBB_VALAXEBP3,QBB_VALAXEBP4,'+
            'QBB_VALAXEBP5,QBB_VALAXEBP6,QBB_VALAXEBP7,QBB_VALAXEBP8,'+
            'QBB_VALAXEBP9,QBB_VALAXEBP10,QBB_VALAXEBP11,QBB_VALAXEBP12,'+
            'QBB_VALAXEBP13,QBB_VALAXEBP14,QBB_VALAXEBP15,QBB_VALAXEBP16,'+
            'QBB_VALAXEBP17,QBB_VALAXEBP18,QBB_VALAXEBP19,QBB_VALAXEBP20,'+
            'QBB_VALAXEBP21,QBB_VALAXEBP22,QBB_VALAXEBP23,QBB_VALAXEBP24,'+
            'QBB_VALAXEBP25,QBB_VALAXEBP26,QBB_VALAXEBP27,QBB_VALAXEBP28,'+
            'QBB_VALAXEBP29,QBB_VALAXEBP30,QBB_DEVISE,QBB_CODEPREVISION,'+
            'QBB_QTEBPC1,QBB_QTEBPC2,QBB_QTEBPC3,QBB_QTEBPC4,QBB_QTEBPC5,'+
            'QBB_QTEBPC6,QBB_QTEBPC7,QBB_QTEBPC8,QBB_QTEBPC9,QBB_QTEBPC10,'+
            'QBB_QTEBPC11,QBB_QTEBPC12,QBB_QTEBPC13,QBB_QTEBPC14,QBB_QTEBPC15,'+
            'QBB_QTEBPC16,QBB_QTEBPC17,QBB_QTEBPC18,QBB_QTEBPC19,QBB_QTEBPC20,QBB_CAC1,'+
            'QBB_DATEFIXATION) '+
            'SELECT "'+codeSession+'","'+'GLOBAL'+'",QBV_CODEIDENT,QBV_MODEGESTPREV,'+
            'QBV_NATURECMD,QBV_CODEMAILLE,QBV_DATEPIVOT,QBV_EXTRAPOLABLE,'+
            'QBV_QTET1*QBR_COEFFRETENU,QBV_QTET2*QBR_COEFFRETENU,'+
            'QBV_QTET3*QBR_COEFFRETENU,QBV_QTET4*QBR_COEFFRETENU,'+
            'QBV_QTET5*QBR_COEFFRETENU,QBV_QTET6*QBR_COEFFRETENU,'+
            'QBV_QTET7*QBR_COEFFRETENU,QBV_QTET8*QBR_COEFFRETENU,'+
            'QBV_QTET9*QBR_COEFFRETENU,QBV_QTET10*QBR_COEFFRETENU,'+
            'QBV_QTET11*QBR_COEFFRETENU,QBV_QTET12*QBR_COEFFRETENU,'+
            'QBV_QTET13*QBR_COEFFRETENU,QBV_QTET14*QBR_COEFFRETENU,'+
            'QBV_QTET15*QBR_COEFFRETENU,QBV_QTET16*QBR_COEFFRETENU,'+
            'QBV_QTET17*QBR_COEFFRETENU,QBV_QTET18*QBR_COEFFRETENU,'+
            'QBV_QTET19*QBR_COEFFRETENU,QBV_QTET20*QBR_COEFFRETENU,'+
            'QBV_QTE1*QBR_COEFFRETENU,QBV_QTE2,QBV_QTE3,QBV_QTE4,QBV_QTE5,QBV_QTE6,'+
            'QBV_VALAXEP1,QBV_VALAXEP2,QBV_VALAXEP3,QBV_VALAXEP4,'+
            'QBV_VALAXEP5,QBV_VALAXEP6,QBV_VALAXEP7,'+codeValAxe8+','+
            'QBV_VALAXEP9,QBV_VALAXEP10,QBV_VALAXEP11,QBV_VALAXEP12,'+
            'QBV_VALAXEP13,QBV_VALAXEP14,QBV_VALAXEP15,QBV_VALAXEP16,'+
            'QBV_VALAXEP17,QBV_VALAXEP18,QBV_VALAXEP19,QBV_VALAXEP20,'+
            'QBV_VALAXEP21,QBV_VALAXEP22,QBV_VALAXEP23,QBV_VALAXEP24,'+
            'QBV_VALAXEP25,QBV_VALAXEP26,QBV_VALAXEP27,QBV_VALAXEP28,'+
            'QBV_VALAXEP29,QBV_VALAXEP30,QBV_DEVISE,QBV_CODEPREVISION,'+
            'QBV_QTET1,QBV_QTET2,QBV_QTET3,QBV_QTET4,QBV_QTET5,'+
            'QBV_QTET6,QBV_QTET7,QBV_QTET8,QBV_QTET9,QBV_QTET10,'+
            'QBV_QTET11,QBV_QTET12,QBV_QTET13,QBV_QTET14,QBV_QTET15,'+
            'QBV_QTET16,QBV_QTET17,QBV_QTET18,QBV_QTET19,QBV_QTET20,QBV_QTE1,QBV_DATEFIXATION '+
            'FROM QBPPIVOT,QBPARBRE WHERE '+codeWhere+
            ' AND QBR_CODESESSION="'+codeSession+'"'+code+codeRestriction,
            'BPrevision (RemplitQBPBudgetprevAvecCoeff).');
end;


procedure ModifQBPBudgetprevCoeff(const codeSession:hString;
          coeff:double;
          TabValeurAxe:array of hString);
var DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin:TDateTime;
    SaiCmd,SaiCmdRef:hString;
    code,iSt:hString;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    nivMax,j:integer;
begin
 if not existeSql('SELECT QBB_CODESESSION FROM QBPBUDGETPREV '+
              'WHERE QBB_CODESESSION="'+codeSession+'"')
  then exit;

 //recupère les données concernant la session
 ChercheDateDDateFPeriode(codeSession,DatePerCDeb,DatePerCFin,DatePerRefDeb,DatePerRefFin);
 ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 //cherche les codes sql
 //on prend les codes jusqu'à nivmax définit
 //pour qbpbudgetprev
 code:='';
 for j:=1 to NivMax+1 do
  begin
   ist:=IntToStr(j);
   if TabNumValAxe[j]<>''
    then code:=code+' AND QBB_VALAXEBP'+TabNumValAxe[j]+'="'+TabValeurAxe[j]+'" ';
  end;

 //update dans qbpbudgetprev
 if SaiCmdRef =''
  then MExecuteSql('UPDATE QBPBUDGETPREV '+
            'SET QBB_QTEBPT1=QBB_QTEBPT1*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT2=QBB_QTEBPT2*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT3=QBB_QTEBPT3*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT4=QBB_QTEBPT4*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT5=QBB_QTEBPT5*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT6=QBB_QTEBPT6*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT7=QBB_QTEBPT7*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT8=QBB_QTEBPT8*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT9=QBB_QTEBPT9*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT10=QBB_QTEBPT10*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT11=QBB_QTEBPT11*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT12=QBB_QTEBPT12*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT13=QBB_QTEBPT13*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT14=QBB_QTEBPT14*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT15=QBB_QTEBPT15*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT16=QBB_QTEBPT16*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT17=QBB_QTEBPT17*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT18=QBB_QTEBPT18*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT19=QBB_QTEBPT19*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT20=QBB_QTEBPT20*'+strFPoint4(coeff)+','+
            'QBB_QTEBP2=QBB_QTEBP2*'+strFPoint4(coeff)+','+
            'QBB_QTEBP3=QBB_QTEBP3*'+strFPoint4(coeff)+','+
            'QBB_QTEBP4=QBB_QTEBP4*'+strFPoint4(coeff)+','+
            'QBB_QTEBP5=QBB_QTEBP5*'+strFPoint4(coeff)+' '+
            'WHERE QBB_CODESESSION="'+codeSession+
            '" AND QBB_DATEBP>="'+USDATETIME(DatePerCDeb)+
            '" AND QBB_DATEBP<="'+USDATETIME(DatePerCFin)+'" '+code,
            'BPrevision (ModifQBPBudgetprevCoeff).')
  else MExecuteSql('UPDATE QBPBUDGETPREV '+
            'SET QBB_QTEBPT1=QBB_QTEBPT1*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT2=QBB_QTEBPT2*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT3=QBB_QTEBPT3*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT4=QBB_QTEBPT4*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT5=QBB_QTEBPT5*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT6=QBB_QTEBPT6*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT7=QBB_QTEBPT7*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT8=QBB_QTEBPT8*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT9=QBB_QTEBPT9*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT10=QBB_QTEBPT10*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT11=QBB_QTEBPT11*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT12=QBB_QTEBPT12*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT13=QBB_QTEBPT13*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT14=QBB_QTEBPT14*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT15=QBB_QTEBPT15*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT16=QBB_QTEBPT16*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT17=QBB_QTEBPT17*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT18=QBB_QTEBPT18*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT19=QBB_QTEBPT19*'+strFPoint4(coeff)+','+
            'QBB_QTEBPT20=QBB_QTEBPT20*'+strFPoint4(coeff)+','+
            'QBB_QTEBP2=QBB_QTEBP2*'+strFPoint4(coeff)+','+
            'QBB_QTEBP3=QBB_QTEBP3*'+strFPoint4(coeff)+','+
            'QBB_QTEBP4=QBB_QTEBP4*'+strFPoint4(coeff)+','+
            'QBB_QTEBP5=QBB_QTEBP5*'+strFPoint4(coeff)+' '+
            'WHERE QBB_CODESESSION="'+codeSession+
            '" AND QBB_VALAXEBP8="'+SaiCmd+'" '+code,
            'BPrevision (ModifQBPBudgetprevCoeff).');
end;

end.
