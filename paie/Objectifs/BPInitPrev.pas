unit BPInitPrev;

interface

uses HEnt1;

procedure BPIntialisePrevision(const codeSession:hString);

implementation

uses Sysutils,HCtrls,
     {$IFNDEF EAGLCLIENT} {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$ELSE} Utob,{$ENDIF}
     CstCommun,Uutil,BPFctSession,BPBasic,BPFctArbre,BPUtil;

function DonneCodeSqlPourSaisiePrevPGI(niveauI:integer;
         TabCodeAxe:array of hString;
         DateRefDeb,DateRefFin:TDateTime;
         codeRestriction,CodeJoinR:hString;
         var champs,joins:hString;
         var histoval,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double; const CodeSession: hString):hString;
var Q:TQuery;
    Chps,Table,Where,join,champsGB:hString;
begin
  champs:='';
  joins:='';
  Table:='';
  Where:='';
  join:='';

  Chps:=' SUM(GL_QTEFACT),SUM(GL_TOTALTTC),SUM(GL_TOTALHT),'+
        'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
        'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
        'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
        'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)) ';
  Table:='LIGNE';
  Where:=' GL_DATEPIECE>="'+USDATETIME(DateRefDeb)+
         '" AND GL_DATEPIECE<="'+USDATETIME(DateRefFin)+'" ';

  if niveauI=1 then
  begin
    Q:=MOpenSql('SELECT '+Chps+
                ' FROM '+Table+' '+join+CodeJoinR+
                ' WHERE '+Where+' '+codeRestriction,
                'BPInitPrev (DonneCodeSqlPourSaisiePrevPGI).',true);
    if not Q.eof then
    begin
      QteVal:=valeur(Q.fields[0].asString);
      histoVal:=valeur(Q.fields[1].asString);
      CAVal2:=valeur(Q.fields[2].asString);
      CAVal3:=valeur(Q.fields[3].asString);
      CAVal4:=valeur(Q.fields[4].asString);
      CAVal5:=valeur(Q.fields[5].asString);
      CAVal6:=valeur(Q.fields[6].asString);
    end;
    ferme(Q);

  end;

  DonneCodeChpJoinPGINbNivLib(TabCodeAxe,niveauI,'',champs,champsGB,joins,CodeSession);

  result:='SELECT '+Chps+','+champs+
          ' FROM '+Table+' '+joins+CodeJoinR+
          ' WHERE '+Where+' '+codeRestriction+
          ' GROUP BY '+champsGB;
end;


function DonneCodeSqlPourSaisiePrev(const CodeSession: hString; NivDelai:boolean;niveauI,NbQuinz:integer;
         DateRefDeb,DateRefFin,DateCDeb,DateCfin:TDateTime;
         TabCodeAxe:array of hString;
         TabNumValAxe:array of hString;
         SaiCmdRef,SaiCmd,structure,codeRestriction,codeRestrictionP,ChpPrev,BpInitialise:hString;
         var cont:boolean):hString;
var champs,champsGB,champs2,champs3,ChpVirgule,ChpSep:hString;
    codeWhereDate:hString;
    DateWdeb,DateWfin: Tdatetime;
    i,niv:integer;
    code,codeWhereQuinz,SaiW:hString;
begin
 champs:='';
 champs2:='';
 champs3:='';


 if NiveauI=1
  then
   begin
    champs:='QBV_VALAXEP'+TabNumValAxe[1];
    champs2:='QBV_VALAXEP'+TabNumValAxe[1];
    champs3:='QBV_VALAXEP'+TabNumValAxe[1];
   end
  else
   begin
    Niv:=niveauI;
    if NivDelai
     then Niv:=niveauI-1;
    for i:=1 to niv do
     begin
      if TabNumValAxe[i]=''
       then cont:=false;
      ChpVirgule:='';
      ChpSep:='';
      if champs<>''
       then
        begin
         ChpVirgule:=',';
         ChpSep:='||';
        end;
      if ChpPrev=''
       then
        begin
         champs:=champs+ChpVirgule+'QBV_VALAXEP'+TabNumValAxe[i];
         champs2:=champs2+ChpSep+'QBV_VALAXEP'+TabNumValAxe[i];
         champs3:=champs3+ChpSep+'P.QBV_VALAXEP'+TabNumValAxe[i];
        end
       else
        begin
         if i>=2
          then
           begin
            champs:=champs+ChpVirgule+'QBV_VALAXEP'+TabNumValAxe[i-1];
            champs2:=champs2+ChpSep+'QBV_VALAXEP'+TabNumValAxe[i-1];
            champs3:=champs3+ChpSep+'P.QBV_VALAXEP'+TabNumValAxe[i-1];
           end;
        end;
      end;
    end;

 if ChpPrev<>''
  then
   begin
    champs:=champs+','+ChpPrev;
    champs2:=champs2+'||'+ChpPrev;
    champs3:=champs3+'||P.'+ChpPrev;
   end;
{
 if OkSessionPermanent(codeSession) then
 begin
   DateWdeb:= DateRefDeb;
   DateWfin:= DateRefFin;
   SaiW:=SaiCmdRef;
 end
   else
 begin
   DateWdeb:= DateCDeb;
   DateWfin:= DateCFin;
   SaiW:=SaiCmd;
 end;
}
   DateWdeb:= DateCDeb;
   DateWfin:= DateCFin;
   SaiW:=SaiCmd;

  codeWhereDate:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateWDeb)+
                  '" AND QBV_DATEPIVOT<="'+USDATETIME(DateWFin)+'" ';
 if SaiW<>''
  then codeWhereDate:=' AND QBV_VALAXEP8="'+SaiW+'" ';

 code:='';
 codeWhereQuinz:='';
 champsGB:=champs;
 if NivDelai
  then
   begin
    champs:=champs+'," "';
    champs2:=champs2+'||" "';
    champs3:=champs3+'||" "';

    case BPInitialise[1] of
     '1':code:=',QBV_DATEPIVOT ';
     '2':code:=',WEEK(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';
     '3':begin
          code:=',MONTH(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';
          if NbQuinz=1
           then codeWhereQuinz:=' AND DAY(QBV_DATEPIVOT)<=15'
           else codeWhereQuinz:=' AND DAY(QBV_DATEPIVOT)>15';
         end;
     '4':code:=',MONTH(QBV_DATEPIVOT),YEAR(QBV_DATEPIVOT) ';
     else code:=',QBV_DATEPIVOT ';
    end;
   end;

 result:='SELECT SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
         'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+'+
         'QBV_QTET10+QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+'+
         'QBV_QTET15+QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+'+
         'QBV_QTET20),SUM(QBV_QTE1),SUM(QBV_QTE2),SUM(QBV_QTE3),'+
         'SUM(QBV_QTE4),SUM(QBV_QTE5),SUM(QBV_QTE6),'+
         'SUM(QBV_QTET1),SUM(QBV_QTET2),SUM(QBV_QTET3),SUM(QBV_QTET4),'+
         'SUM(QBV_QTET5),SUM(QBV_QTET6),SUM(QBV_QTET7),SUM(QBV_QTET8),'+
         'SUM(QBV_QTET9),SUM(QBV_QTET10),SUM(QBV_QTET11),SUM(QBV_QTET12),'+
         'SUM(QBV_QTET13),SUM(QBV_QTET14),SUM(QBV_QTET15),SUM(QBV_QTET16),'+
         'SUM(QBV_QTET17),SUM(QBV_QTET18),SUM(QBV_QTET19),SUM(QBV_QTET20), '+
         champs+code+
         ',QBV_DEVISE FROM QBPPIVOT WHERE QBV_CODESTRUCT="'+structure+
         '" AND '+codeWhereDate+codeRestriction+codeWhereQuinz+
         ' GROUP BY '+champsGB+code+',QBV_DEVISE ';
end;


procedure ChercheNoeudPereNivXPrev(niveau:integer;CodeSession,devise:hString;
          TabValeur:array of hString;
          var numnoeudPere:hString;
          var histoVal,QteVal,CA2Val,CA3Val,CA4Val,CA5Val,CA6Val,CoeffRetenu:double;
          var OkPere:boolean);
var noeudP:hString;
    Q:TQuery;
begin
 if niveau=1
  then
   begin
    ArbreTotalNivRef(codeSession,'0',QteVal,histoVal,CA2Val,CA3Val,CA4Val,CA5Val,CA6Val);
    exit;
   end;
 histoVal:=0;
 //cherche noeudPere
 if niveau=2
  then noeudP:='0';
 if niveau>2
  then ChercheNoeudPereNivXPrev(niveau-1,CodeSession,devise,
                            TabValeur,noeudP,histoVal,QteVal,
                            CA2Val,CA3Val,CA4Val,CA5Val,CA6Val,CoeffRetenu,OkPere);

 //on cherche dans la table qbparbre le numéro de noeud
 //correspondant à la session, ayant pour valeuraxe la valeur donnée
 Q:=MOpenSql('SELECT QBR_NUMNOEUD,QBR_REF1,QBR_QTEREF,'+
             'QBR_REF2,QBR_REF3,QBR_REF4,QBR_REF5,QBR_REF6,QBR_COEFFRETENU '+
             'FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+CodeSession+
             '" AND QBR_VALEURAXE="'+TabValeur[niveau-2]+
             '" AND QBR_DEVISE="'+devise+
             '" AND QBR_NUMNOEUDPERE="'+noeudP+'"'
             ,'BPInitPrev (ChercheNoeudPereNivXPrev).',true);
 if not Q.eof
  then
   begin
    numnoeudPere:=Q.fields[0].asString;
    histoVal:=Q.fields[1].asFloat;
    QteVal:=Q.fields[2].asFloat;
    CA2Val:=Q.fields[3].asFloat;
    CA3Val:=Q.fields[4].asFloat;
    CA4Val:=Q.fields[5].asFloat;
    CA5Val:=Q.fields[6].asFloat;
    CA6Val:=Q.fields[7].asFloat;
    CoeffRetenu:=Q.fields[8].asFloat;
   end
  else OkPere:=false;
 ferme(Q);
end;


function ChercheCodeValeurAxeDelai(const Val1,Val2:integer;BPInitialise:hString):TDateTime;
var DelaiMaille:hString;
    code,DateI:TDateTime;
begin
 code:=0;
 case BPInitialise[1] of
  '2':code:=PremierJourSemaine(Val1,Val2);
  '3':code:=EncodeDate(val2,val1,1);
  '4':begin
       DelaiMaille:=DonneParamS(ps_BPDelaiMaille);
       DateI:=EncodeDate(val2,val1,1);
       case DelaiMaille[1] of
        '1':code:=DEBUTDEMOIS(DateI);
        '2':code:=DEBUTDEMOIS(DateI)+14;
        '3':code:=FINDEMOIS(DateI);
       end;
      end;
  '6':code:=EncodeDate(val2,val1,15);
 end;
 result:=code;
end;


//insertion niveau x
procedure InsertIntoQbpSaisiePrevNivXPrev(NivDelai:boolean;nivMaxPrev:integer;const niveau,structure,
          codesession,codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef:hString;
          BpInitialise,CodeAxePrev,ChpPrev:hString;
          DateRefDeb,DateRefFin,DateCDeb,DateCFin:TDateTime;
          TabCodeAxe:array of hString;
          TabNumValAxe:array of hString;
          nivMax:integer;NbQuinz:integer);
var Q,Q2:TQuery;
    numnoeud,numnoeudpere,devise,libelleValeur:hString;
    champs,joins,ValeurSt,codeSql,ChpPrevX:hString;
    niveauI,i:integer;
    cont,OkPere,OkSessionCalculParTaille:boolean;
    TabValeur:array [1..11] of hString;
    TabQte:array [0..20] of double;
    valeurHisto,histoVal,histoprct:double;
    QteVal,QtePrct:double;
    CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,CoeffRetenu:double;
    Qte,CA2,CA3,CA4,CA5,CA6:double;
    CAVal2Prct,CAVal3Prct,CAVal4Prct,CAVal5Prct,CAVal6Prct:double;
    coeff,histo,prevu,realise,PrevuArrondi,errPrevu:double;
    totVal,TotQteVal,TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5:double;
    ValDelai,ANDelai:integer;
    codeDelai:TDateTime;
    rapport:double;
    ListeQte,ListeQteRes: array[0..19] of double;
begin
  niveauI:=VALEURI(niveau);
  cont:=true;
  numnoeudpere:='0';
  if not BPOkOrli then
  begin
    codeSql:=DonneCodeSqlPourSaisiePrevPGI(niveauI,TabCodeAxe,
                                     DateRefDeb,DateRefFin,codeRestriction,CodeJoinR,
                                     champs,joins,histoval,QteVal,
                                     CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,CodeSession);
  end    else
  //-----------------> ORLI
    codeSql:=DonneCodeSqlPourSaisiePrev(CodeSession,NivDelai,niveauI,NbQuinz,DateRefDeb,DateRefFin,
                                     DateCDeb,DateCFin,
                                     TabCodeAxe,TabNumValAxe,SaiCmdRef,SaiCmd,
                                     structure,codeRestriction,'',ChpPrev,BpInitialise,
                                     cont);
  //ORLI <-----------------
  if ChpPrev<>'' then cont:=true;
  if cont then
  begin
    numnoeud:=IntToStr(BPIncrementeNumNoeud(codesession));
    for i:=1 to 11 do
     TabValeur[i]:='';
    ChpPrevX:='';
    errPrevu:=0;
    okSessionCalculParTaille:=SessionCalculParTaille(codeSession);
    Q:=MOpenSql(CodeSql,'BPInitPrev (InsertIntoQbpSaisiePrevNivXPrev).',true);
    while not Q.eof do
    begin
      libelleValeur:='';
      codeDelai:=0;

      Qte:=Q.fields[0].asFloat;
      valeurHisto:=Q.fields[1].asFloat;
      if not BPOkOrli then
      begin
        CA2:=Q.fields[2].asFloat;
        CA3:=Q.fields[3].asFloat;
        CA4:=Q.fields[4].asFloat;
        CA5:=Q.fields[5].asFloat;
        CA6:=Q.fields[6].asFloat;
        for i:=1 to niveauI do
          TabValeur[i]:=Q.fields[i+6].asString;
        devise:=Q.fields[niveauI+2+6].asString;
        libelleValeur:=Q.fields[niveauI+1+6].asString;
      end
       else
      //-----------------> ORLI
      begin
        CA2:=Q.fields[2].asFloat;
        CA3:=Q.fields[3].asFloat;
        CA4:=Q.fields[4].asFloat;
        CA5:=Q.fields[5].asFloat;
        CA6:=Q.fields[6].asFloat;
        for i:=1 to niveauI do
          TabValeur[i]:=Q.fields[i+26].asString;
        if NivDelai then
        begin
          if BPInitialise='5' then
          begin
            codeDelai:=(FindPeriodeP(Q.fields[niveauI+1+26].asDateTime)).datedeb;
            devise:=Q.fields[niveauI+2+26].asString;
          end
             else
          begin
            ValDelai:=Q.fields[niveauI+1+26].asInteger;
            ANDelai:=Q.fields[niveauI+2+26].asInteger; ///  +1;   Pourquoi +1 ????
            codeDelai:=ChercheCodeValeurAxeDelai(ValDelai,ANDelai,BPInitialise);
            devise:=Q.fields[niveauI+3+26].asString;
          end;
        end
          else devise:=Q.fields[niveauI+1+26].asString;
      end;
      //ORLI <-----------------
      numnoeud:=IntToStr(VALEURI(numnoeud)+1);
      //cherche le noeud pere
      OkPere:=true;
      ChercheNoeudPereNivXPrev(niveauI,CodeSession,devise,TabValeur,numnoeudPere,histoVal,
                           QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,
                           CoeffRetenu,OkPere);
   //   if (niveauI>2) and (numnoeudPere='0')
   //    then OkPere:=false;

      prevu:=Qte*coeffretenu;
      PrevuArrondi:=round(prevu+errPrevu);
      errPrevu:=errPrevu+(prevu-PrevuArrondi);
      prevu:=PrevuArrondi;
      if (VALEURI(niveau)=NivMaxPrev) and okSessionCalculParTaille and NivDelai
       then
      begin
        rapport:=0;
        if qte<>0 then rapport:=prevu/qte;
        for i:=7 to 26 do
          ListeQte[i-7]:= valeur(Q.fields[i].asString) ;
        BPCalculQteListProrata(0,ListeQte,Rapport,ListeQteRes);
      end
       else
      begin
        for i:=0 to 19 do
          ListeQteRes[i]:=0;
      end;
      if (not NivDelai) and (ChpPrev='')
        then CoeffRetenu:=1;
      histo:=QteVal;
      coeff:=1;
      realise:=Qte;
      //Qte
      if QteVal=0  then QteVal:=Qte;
      if QteVal<>0 then QtePrct:=(Qte/QteVal)*100
                   else QtePrct:=0;
      //CA1
      if histoVal=0  then histoVal:=valeurHisto;
      if histoVal<>0 then histoprct:=(valeurHisto/histoVal)*100
                     else histoprct:=0;
      //CA2
      if CAVal2=0  then CAVal2:=CA2;
      if CAVal2<>0 then CAVal2Prct:=(CA2/CAVal2)*100
                   else CAVal2Prct:=0;
      //CA3
      if CAVal3=0  then CAVal3:=CA3;
      if CAVal3<>0 then CAVal3Prct:=(CA3/CAVal3)*100
                   else CAVal3Prct:=0;
      //CA4
      if CAVal4=0  then CAVal4:=CA4;
      if CAVal4<>0 then CAVal4Prct:=(CA4/CAVal4)*100
                   else CAVal4Prct:=0;
      //CA5
      if CAVal5=0  then CAVal5:=CA5;
      if CAVal5<>0 then CAVal5Prct:=(CA5/CAVal5)*100
                   else CAVal5Prct:=0;
      //CA6
      if CAVal6=0  then CAVal6:=CA6;
      if CAVal6<>0 then CAVal6Prct:=(CA6/CAVal6)*100
                   else CAVal6Prct:=0;
      if NivDelai
        then ValeurSt:=DateTimeToStr(codeDelai)
        else ValeurSt:=TabValeur[niveauI];
      TabValeur[niveauI]:='';
      if (BPOkOrli) and (CodeAxePrev<>'')
       //-----------------> ORLI
       then TabCodeAxe[niveauI]:=CodeAxeprev;
       //ORLI <-----------------
      //coeffretenu:=coeff;
      if OkPere
       then MExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,'+
                 'QBR_NUMNOEUDPERE,QBR_VALEURAXE,QBR_REF1,QBR_REFPRCT1,'+
                 'QBR_OP1,QBR_OPPRCT1, '+
                 'QBR_NIVEAU,QBR_CODEAXE,QBR_VALAXENIV1,QBR_VALAXENIV2,'+
                 'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
                 'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_DEVISE,'+
                 'QBR_LIBVALAXE,QBR_QTEREFPRCT,QBR_QTEREF,QBR_QTECPRCT,QBR_QTEC,'+
                 'QBR_REFPRCT2,QBR_REF2,QBR_OPPRCT2,QBR_OP2,'+
                 'QBR_REFPRCT3,QBR_REF3,QBR_OPPRCT3,QBR_OP3,'+
                 'QBR_REFPRCT4,QBR_REF4,QBR_OPPRCT4,QBR_OP4,'+
                 'QBR_REFPRCT5,QBR_REF5,QBR_OPPRCT5,QBR_OP5,'+    
                 'QBR_REFPRCT6,QBR_REF6,QBR_OPPRCT6,QBR_OP6,'+
                 'QBR_COEFFCALCUL,QBR_HISTO,QBR_PREVU,QBR_REALISE,QBR_COEFFRETENU,'+
                 'QBR_VALBLOQUE,QBR_VALBLOQUETMP,QBR_DATEDELAI,QBR_QTERETENUE,'+
                 'QBR_QTET1,QBR_QTET2,QBR_QTET3,QBR_QTET4,QBR_QTET5,'+
                 'QBR_QTET6,QBR_QTET7,QBR_QTET8,QBR_QTET9,QBR_QTET10,'+
                 'QBR_QTET11,QBR_QTET12,QBR_QTET13,QBR_QTET14,QBR_QTET15,'+
                 'QBR_QTET16,QBR_QTET17,QBR_QTET18,QBR_QTET19,QBR_QTET20) '+
                 'VALUES ("'+codesession+'","'+numnoeud+
                 '","'+numnoeudPere+'","'+ValeurSt+
                 '","'+strFPoint4(valeurHisto)+
                 '","'+strFPoint4(histoprct)+
                 '","'+strFPoint4(valeurHisto*coeffretenu)+
                 '","'+strFPoint4(histoprct)+
                 '","'+niveau+'","'+TabCodeAxe[niveauI]+
                 '","'+TabValeur[1]+'","'+TabValeur[2]+'","'+TabValeur[3]+
                 '","'+TabValeur[4]+'","'+TabValeur[5]+'","'+TabValeur[6]+
                 '","'+TabValeur[7]+'","'+TabValeur[8]+'","'+TabValeur[9]+
                 '","'+Devise+'","'+RemplaceCotteEspacePourSql(libelleValeur)+
                 '","'+strFPoint4(QtePrct)+'","'+strFPoint4(Qte)+
                 '","'+strFPoint4(QtePrct)+'","'+strFPoint4(Qte*coeffretenu)+
                 '","'+strFPoint4(CAVal2Prct)+'","'+strFPoint4(CA2)+
                 '","'+strFPoint4(CAVal2Prct)+'","'+strFPoint4(CA2*coeffretenu)+
                 '","'+strFPoint4(CAVal3Prct)+'","'+strFPoint4(CA3)+
                 '","'+strFPoint4(CAVal3Prct)+'","'+strFPoint4(CA3*coeffretenu)+
                 '","'+strFPoint4(CAVal4Prct)+'","'+strFPoint4(CA4)+
                 '","'+strFPoint4(CAVal4Prct)+'","'+strFPoint4(CA4*coeffretenu)+
                 '","'+strFPoint4(CAVal5Prct)+'","'+strFPoint4(CA5)+
                 '","'+strFPoint4(CAVal5Prct)+'","'+strFPoint4(CA5*coeffretenu)+
                 '","'+strFPoint4(CAVal6Prct)+'","'+strFPoint4(CA6)+
                 '","'+strFPoint4(CAVal6Prct)+'","'+strFPoint4(CA6*coeffretenu)+
                 '","'+strFPoint4(Coeff)+'","'+strFPoint4(Histo)+
                 '","'+strFPoint4(Prevu)+'","'+strFPoint4(Realise)+
                 '","'+strFPoint4(coeffretenu)+'","-","-","'+USDATETIME(codeDelai)+
                 '","'+strFPoint4(prevu)+
                 '",'+strFpoint(ListeQteRes[0])+','+strFpoint(ListeQteRes[1])+
                 ','+strFpoint(ListeQteRes[2])+','+strFpoint(ListeQteRes[3])+
                 ','+strFpoint(ListeQteRes[4])+','+strFpoint(ListeQteRes[5])+
                 ','+strFpoint(ListeQteRes[6])+','+strFpoint(ListeQteRes[7])+
                 ','+strFpoint(ListeQteRes[8])+','+strFpoint(ListeQteRes[9])+
                 ','+strFpoint(ListeQteRes[10])+','+strFpoint(ListeQteRes[11])+
                 ','+strFpoint(ListeQteRes[12])+','+strFpoint(ListeQteRes[13])+
                 ','+strFpoint(ListeQteRes[14])+','+strFpoint(ListeQteRes[15])+
                 ','+strFpoint(ListeQteRes[16])+','+strFpoint(ListeQteRes[17])+
                 ','+strFpoint(ListeQteRes[18])+','+strFpoint(ListeQteRes[19])+
                 ')',
                 'BPInitPrev (InsertIntoQbpSaisiePrevNivxPrev).');
      Q.next;
    end;
    ferme(Q);
  end;
end;

procedure BPIntialisePrevision(const codeSession:hString);
var nivMax,niv,nivMaxPrev:integer;
    structure,codeRestriction,BPInitialise,codeMag,SaiCmd,SaiCmdRef,codeJoinR:hString;
    TabNumValAxe,TabCodeAxe:array [1..11] of hString;
    DateDebCourante,DateFinCourante,DateDebReference,DateFinReference:TDateTime;
begin
 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);
 //cherche niv max de la session de prev
 //niv article
 nivMaxPrev:=nivMax+1;
 //niv coloris
 if GestionBPColoris
  then nivMaxPrev:=nivMaxPrev+1;
 //niv FS
 if GestionBPFS
  then nivMaxPrev:=nivMaxPrev+1;
 //niv Magasin
 if GestionBPMagasin
  then nivMaxPrev:=nivMaxPrev+1;
 //niv delai
 if (DonneMethodeSession(codeSession)='1')
  then nivMaxPrev:=nivMaxPrev+1;

 structure:=ChercheSessionStructure(codeSession);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);
                                        
 codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,codeJoinR);
 ChercheDateDDateFPeriode(codeSession,DateDebCourante,DateFinCourante,
                          DateDebReference,DateFinReference);
 ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
 SessionCalculParDelai(codeSession,BPInitialise);

 MExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NIVEAU>'+STRFPOINT(nivMax),'BPInitPrev (InsertIntoQbpSaisiePrevNivxPrev).');

 //pour axe article
 niv:=nivMax+1;
 TabCodeAxe[niv+1]:='ARTICLE';
 InsertIntoQbpSaisiePrevNivXPrev(false,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                 codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                 'ARTICLE','QBV_VALAXEP2',
                                 DateDebReference,DateFinReference,
                                 DateDebCourante,DateFinCourante,TabCodeAxe,
                                 TabNumValAxe,nivMax,0);
 TabNumValAxe[niv+1]:='2';
 //pour axe coloris
 //si gestion des coloris
 if GestionBPColoris
  then
   begin
    niv:=niv+1;
    TabCodeAxe[niv+1]:='COLORIS';
    InsertIntoQbpSaisiePrevNivXPrev(false,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                    codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                    'COLORIS','QBV_VALAXEP3',
                                    DateDebReference,DateFinReference,
                                    DateDebCourante,DateFinCourante,TabCodeAxe,
                                    TabNumValAxe,nivMax,0);
    TabNumValAxe[niv+1]:='3';
   end;

 //pour axe fs
 //si gestion des fs
 if GestionBPFs
  then
   begin
    niv:=niv+1;
    TabCodeAxe[niv+1]:='FS';
    InsertIntoQbpSaisiePrevNivXPrev(false,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                    codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                    'FS','QBV_VALAXEP4',
                                    DateDebReference,DateFinReference,
                                    DateDebCourante,DateFinCourante,TabCodeAxe,
                                    TabNumValAxe,nivMax,0);
    TabNumValAxe[niv+1]:='4';
   end;

 //pour axe magasin (depot)
 //si gestion des magasins
 if GestionBPMagasin
  then
   begin
    niv:=niv+1;
    codeMag:='MAGASIN';
    if ContextBP in [0,2,3]
     then codeMag:='DEPOT';
    TabCodeAxe[niv+1]:=codeMag;
    InsertIntoQbpSaisiePrevNivXPrev(false,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                    codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                    codeMag,'QBV_VALAXEP7',
                                    DateDebReference,DateFinReference,
                                    DateDebCourante,DateFinCourante,TabCodeAxe,
                                    TabNumValAxe,nivMax,0);
    TabNumValAxe[niv+1]:='7';
   end;

 //niveau delai
 //si session definie par délai
 if (DonneMethodeSession(codeSession)='1')
  then
   begin
    niv:=niv+1;
    TabCodeAxe[niv+1]:='DELAI';
    InsertIntoQbpSaisiePrevNivXPrev(true,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                    codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                    'DELAI','',
                                    DateDebReference,DateFinReference,
                                    DateDebCourante,DateFinCourante,TabCodeAxe,
                                    TabNumValAxe,nivMax,1);
    if BPInitialise='3'
     then InsertIntoQbpSaisiePrevNivXPrev(true,nivMaxPrev,IntToStr(niv),structure,codeSession,
                                    codeRestriction,CodeJoinR,SaiCmd,SaiCmdRef,BPInitialise,
                                    'DELAI','',
                                    DateDebReference,DateFinReference,
                                    DateDebCourante,DateFinCourante,TabCodeAxe,
                                    TabNumValAxe,nivMax,2);
   end;

 MiseAJourPrctArbre(codeSession,false,true);

 MiseAJourQtePrevueQteRetenue(codeSession);

 //mise à jour dans qbpsessionbp session init
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="X",QBS_INITPREVISION="X" '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPInitPrev (BPIntialisePrevision).');
 if (DonneMethodeSession(codeSession)='1')
  then MExecuteSql('UPDATE QBPSESSIONBP SET QBS_INITDELAI="X" '+
                   'WHERE QBS_CODESESSION="'+codeSession+'"',
                   'BPInitPrev (BPIntialisePrevision).');
end;


end.
