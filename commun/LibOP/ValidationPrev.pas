unit ValidationPrev;

interface

uses SysUtils,
     HCtrls,Hent1;

procedure ValidationPrevisionArbre(const codeSession:string);

procedure ValidationBP(const codeSession:string);

implementation


uses uutil,
     BPMaille,BPFctSession;

//******************************************************************
//
//                Validation Prevision
//
//******************************************************************


procedure ValidationPrevisionArbre(const codeSession:string);
var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    ListMaille:TListMaille;
    i,j,nivMax:integer;
    Mi:TMaille;
    ist,istM,types,numOrdre:string;
    CodeQteTotale,CodeQtePrevu,CodeSql,ChpSql,codeSqlWhere,ChpSqlTotal,ChpSqlPrevu:string;
    ChpSqlPrevuBy:string;
    TabCodeAxe,TabNumValAxe:array [1..11] of string;
begin
 //cherche les dates de la session
 ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);

 //maille
 ListMaille:=TListMaille.create();
 SessionCalculParDelai(codeSession,types);
 types:='4';
 InitialiseListeMaille(VALEURI(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

 //codeAxe
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 //ChpSql
 ChpSql:='';
 ChpSqlTotal:='';
 ChpSqlPrevu:='';
 ChpSqlPrevuBy:='';
 for j:=1 to NivMax+1 do
  begin
   ist:=IntToStr(j);
   istM:=IntToStr(j-1);
   if TabNumValAxe[j]<>''
    then
     begin
      if ChpSql=''
       then ChpSql:=' QBV_VALAXEP'+TabNumValAxe[j]+' '
       else ChpSql:=ChpSql+',QBV_VALAXEP'+TabNumValAxe[j]+' ';
      ChpSqlTotal:=ChpSqlTotal+',QBV_VALAXEP'+TabNumValAxe[j]+' AXETOT'+istM+' ';
      ChpSqlPrevu:=ChpSqlPrevu+',QBE_VALAXENIV'+istM+' AXEP'+istM+' ';
      if ChpSqlPrevuBy=''
       then ChpSqlPrevuBy:=' QBE_VALAXENIV'+istM+' '
       else ChpSqlPrevuBy:=ChpSqlPrevuBy+',QBE_VALAXENIV'+istM+' ';
      codeSqlWhere:=codeSqlWhere+' AND QBV_VALAXEP'+TabNumValAxe[j]+
                    '=AXEP'+istM+' AND QBV_VALAXEP'+TabNumValAxe[j]+
                    '=AXETOT'+istM+' ';
     end;
  end;
 //nivmax
 istM:=IntToStr(NivMax+1);
 if ChpSql=''
  then ChpSql:=' QBV_VALAXEP2 '
  else ChpSql:=ChpSql+',QBV_VALAXEP2 ';
 ChpSqlTotal:=ChpSqlTotal+',QBV_VALAXEP2 AXETOT'+istM+' ';
 ChpSqlPrevu:=ChpSqlPrevu+',QBE_VALAXENIV'+istM+' AXEP'+istM+' ';
 if ChpSqlPrevuBy=''
  then ChpSqlPrevuBy:=' QBE_VALAXENIV'+istM+' '
  else ChpSqlPrevuBy:=ChpSqlPrevuBy+',QBE_VALAXENIV'+istM+' ';
 codeSqlWhere:=codeSqlWhere+' AND QBV_VALAXEP2'+
               '=AXEP'+istM+' AND QBV_VALAXEP2'+
               '=AXETOT'+istM+' ';


 //pour chaque maille
 //on lance insert dans table qbpbudgetprev
 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);

   //qté totale
   CodeQteTotale:=' (SELECT SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5'+
                  '+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
                  'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
                  'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20) '+
                  ' AS QTETOTALE,SUM(QBV_QTE2) CATOT '+ChpSqlTotal+' FROM QBPPIVOT WHERE '+
                  'QBV_DATEPIVOT>="'+USDATETIME(Mi.DateDebCourante)+
                  '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinCourante)+
                  '" GROUP BY '+ChpSql+') VUETOTALE ';

   //qté prévue
   CodeQtePrevu:=' ,(SELECT SUM(QBE_NEWQTEVAL) QTEPREVU,SUM(QBE_NEWVAL) CAPREVU '+
                 ChpSqlPrevu+
                 ' FROM QBPARBREPREV '+
                 'WHERE QBE_CODESESSION="'+codeSession+
                 '" AND QBE_DATEDELAI="'+USDATETIME(Mi.DateDebCourante)+
                 '" GROUP BY '+ChpSqlPrevuBy+') VUEPREVU ';

   numOrdre:='1';
   //code sql
   codeSql:='INSERT INTO QBPBUDGETPREV '+
            '(QBB_CODESESSION,QBB_NUMORDRE,QBB_CODEIDENT,QBB_MODEGESTPREV,'+
            'QBB_NATURECMD,QBB_CODEMAILLE,QBB_DATEBP,QBB_EXTRAPOLABLE,'+
            'QBB_QTEBPT1,QBB_QTEBPT2,QBB_QTEBPT3,QBB_QTEBPT4,QBB_QTEBPT5,'+
            'QBB_QTEBPT6,QBB_QTEBPT7,QBB_QTEBPT8,QBB_QTEBPT9,QBB_QTEBPT10,'+
            'QBB_QTEBPT11,QBB_QTEBPT12,QBB_QTEBPT13,QBB_QTEBPT14,QBB_QTEBPT15,'+
            'QBB_QTEBPT16,QBB_QTEBPT17,QBB_QTEBPT18,QBB_QTEBPT19,QBB_QTEBPT20,'+
            'QBB_QTEBP2,QBB_QTEBP3,QBB_QTEBP4,QBB_QTEBP5,'+
            'QBB_VALAXEBP1,QBB_VALAXEBP2,QBB_VALAXEBP3,QBB_VALAXEBP4,'+
            'QBB_VALAXEBP5,QBB_VALAXEBP6,QBB_VALAXEBP7,QBB_VALAXEBP8,'+
            'QBB_VALAXEBP9,QBB_VALAXEBP10,QBB_VALAXEBP11,QBB_VALAXEBP12,'+
            'QBB_VALAXEBP13,QBB_VALAXEBP14,QBB_VALAXEBP15,QBB_VALAXEBP16,'+
            'QBB_VALAXEBP17,QBB_VALAXEBP18,QBB_VALAXEBP19,QBB_VALAXEBP20,'+
            'QBB_VALAXEBP21,QBB_VALAXEBP22,QBB_VALAXEBP23,QBB_VALAXEBP24,'+
            'QBB_VALAXEBP25,QBB_VALAXEBP26,QBB_VALAXEBP27,QBB_VALAXEBP28,'+
            'QBB_VALAXEBP29,QBB_VALAXEBP30) '+
            'SELECT "'+codeSession+'","'+numOrdre+'",PGIGUID,QBV_MODEGESTPREV,'+
            'QBV_NATURECMD,QBV_CODEMAILLE,QBV_DATEPIVOT'+
            ',QBV_EXTRAPOLABLE,(QBV_QTET1/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET2/QTETOTALE)*QTEPREVU,(QBV_QTET3/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET4/QTETOTALE)*QTEPREVU,(QBV_QTET5/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET6/QTETOTALE)*QTEPREVU,(QBV_QTET7/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET8/QTETOTALE)*QTEPREVU,(QBV_QTET9/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET10/QTETOTALE)*QTEPREVU,(QBV_QTET11/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET12/QTETOTALE)*QTEPREVU,(QBV_QTET13/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET14/QTETOTALE)*QTEPREVU,(QBV_QTET15/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET16/QTETOTALE)*QTEPREVU,(QBV_QTET17/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET18/QTETOTALE)*QTEPREVU,(QBV_QTET19/QTETOTALE)*QTEPREVU,'+
            '(QBV_QTET20/QTETOTALE)*QTEPREVU,QBV_QTE2,QBV_QTE3,QBV_QTE4,QBV_QTE5, '+
            'QBV_VALAXEP1,QBV_VALAXEP2,QBV_VALAXEP3,QBV_VALAXEP4,QBV_VALAXEP5,'+
            'QBV_VALAXEP6,QBV_VALAXEP7,QBV_VALAXEP8,QBV_VALAXEP9,QBV_VALAXEP10,'+
            'QBV_VALAXEP11,QBV_VALAXEP12,QBV_VALAXEP13,QBV_VALAXEP14,QBV_VALAXEP15,'+
            'QBV_VALAXEP16,QBV_VALAXEP17,QBV_VALAXEP18,QBV_VALAXEP19,'+
            'QBV_VALAXEP20,QBV_VALAXEP21,QBV_VALAXEP22,QBV_VALAXEP23,'+
            'QBV_VALAXEP24,QBV_VALAXEP25,QBV_VALAXEP26,QBV_VALAXEP27,'+
            'QBV_VALAXEP28,QBV_VALAXEP29,QBV_VALAXEP30 FROM '+CodeQteTotale+
            CodeQtePrevu+',QBPPIVOT WHERE '+
            'QBV_DATEPIVOT>="'+USDATETIME(Mi.DateDebCourante)+
            '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinCourante)+
            '" AND QTETOTALE>0 '+codeSqlWhere
            ;

   MExecuteSql(codeSql,'ValidationPrev (ValidationPrevisionArbre).');
  end;
  freeListMaille(ListMaille);
end;


//******************************************************************
//
//                Validation Budget Prévision
//
//******************************************************************


procedure ValidationBP(const codeSession:string);
begin
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONVALIDE="X",'+
             'QBS_DATEVALIDATION="'+USDATETIME(date)+
             '" WHERE QBS_CODESESSION="'+codeSession+'"',
             'ValidationPrev (ValidationBudget).');
end;

end.

