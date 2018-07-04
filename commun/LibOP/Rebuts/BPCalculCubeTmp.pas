unit BPCalculCubeTmp;

interface

uses {$IFNDEF EAGLCLIENT}
     dbtables,
     {$ENDIF}
     HCtrls,Sysutils,UTob,HEnt1,
     UGraph,UBasic,UCtx,
     BPMaille,BPBasic,BPFctSession,BPFctArbre;

procedure RemplitTableQBPCubeTmpPgi(const codeSession:string;
          DateDebCourante,DateFinCourante:TDateTime);
procedure RemplitTableQBPCubeTmpOrli(const codeSession:string;
          DateDebCourante,DateFinCourante:TDateTime);

implementation

procedure RemplitTableQBPCubeTmpPrevuObjDelai(const codeSession,
          codeChpSql,codeFiltre:string;
          TabCodeAxe:array of string);
var nivMax,i:integer;
    codesqlArbre,chpArbreLib,joinArbreLib:string;
begin
 nivMax:=ChercheNivMax(codeSession)+1; //pc delai
 codesqlArbre:='';

 for i:=1 to NivMax-1 do
  begin
   if codesqlArbre=''
    then codesqlArbre:=' QBI_VALAXENIV'+IntToStr(i)
    else codesqlArbre:=codesqlArbre+',QBI_VALAXENIV'+IntToStr(i);
  end;

 ChercheCodeSqlChpAxeLib(nivMax,TabCodeAxe,chpArbreLib,joinArbreLib);

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
            codesqlArbre+chpArbreLib+',QBI_DATEDELAI,'+
            '0,QBI_NEWVAL,0,'+
            '0,QBI_NEWQTEVAL,0,'+
            '0,QBI_NEWCAVAL2,0,'+
            '0,QBI_NEWCAVAL3,0,'+
            '0,QBI_NEWCAVAL4,0,'+
            '0,QBI_NEWCAVAL5,0 '+
            ' FROM QBPSAISIEPREV '+joinArbreLib+
            ' WHERE QBI_CODESESSION="'+codeSession+
            '" AND QBI_NIVEAU="'+IntToStr(NivMax)+'" '+codefiltre,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelai).');
 { MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'",'+
            codesqlArbre+chpArbreLib+',QBI_DATEDELAI,'+
            '0,QBI_NEWVAL,0,'+
            '0,QBI_NEWQTEVAL,0,'+
            '0,QBI_NEWCAVAL2,0,'+
            '0,QBI_NEWCAVAL3,0,'+
            '0,QBI_NEWCAVAL4,0,'+
            '0,QBI_NEWCAVAL5,0 '+
            ' FROM QBPSAISIEPREV '+joinArbreLib+
            ' WHERE QBI_CODESESSION="'+codeSession+
             '" AND QBI_NIVEAU="'+IntToStr(NivMax)+'" '+codeFiltre);}

end;

procedure RemplitTableQBPCubeTmpPrevuObjDelaiMaille(const codeSession,
          codeChpSql:string;Mi:TMaille;
          TabCodeAxe:array of string);
var nivMax,i:integer;
    codesqlArbre,chpArbreLib,joinArbreLib:string;
begin
 nivMax:=ChercheNivMax(codeSession)+1; //pc delai
 codesqlArbre:='';

 for i:=1 to NivMax-1 do
  begin
   if codesqlArbre=''
    then codesqlArbre:=' QBI_VALAXENIV'+IntToStr(i)
    else codesqlArbre:=codesqlArbre+',QBI_VALAXENIV'+IntToStr(i);
  end;

 ChercheCodeSqlChpAxeLib(nivMax,TabCodeAxe,chpArbreLib,joinArbreLib);

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+
            codeChpSql+',QBQ_DATECT,QBQ_CODEMAILLE,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
            codesqlArbre+chpArbreLib+',QBI_DATEDELAI,'+STRFPOINT(Mi.code)+
            ',0,QBI_NEWVAL,0,'+
            '0,QBI_NEWQTEVAL,0,'+
            '0,QBI_NEWCAVAL2,0,'+
            '0,QBI_NEWCAVAL3,0,'+
            '0,QBI_NEWCAVAL4,0,'+
            '0,QBI_NEWCAVAL5,0, '+
            '0,QBI_NEWCAVAL6,0 '+
            ' FROM QBPSAISIEPREV '+joinArbreLib+
            ' WHERE QBI_CODESESSION="'+codeSession+
             '" AND QBI_NIVEAU="'+IntToStr(NivMax)+
             '" AND QBI_DATEDELAI>="'+USDATETIME(Mi.DateDebCourante)+
             '" AND QBI_DATEDELAI<="'+USDATETIME(Mi.DateFinCourante)+'"',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
end;

procedure InsertCubeNewLigneArbre(codeSession:string;TabNumValAxe:array of string;
          niv:integer);
var i:integer;
    codeChpSql,codeChpSqlW,codeChpArbre,codeChpVal:string;
begin
 codeChpSql:='';
 codeChpArbre:='';
 codeChpVal:='';
 for i:=1 to niv do
  begin
   if BpOkOrli
    then codeChpSql:=codeChpSql+',QBQ_VALAXECT'+TabNumValAxe[i]
    else codeChpSql:=codeChpSql+',QBQ_VALAXECT'+IntToStr(i);
  end;
 for i:=1 to niv-1 do
   codeChpArbre:=codeChpArbre+',QBI_VALAXENIV'+IntToStr(i);
 codeChpArbre:=codeChpArbre+',QBI_VALEURAXE,QBI_DATEDELAI ';

 codeChpVal:=',0,QBI_NEWVAL,0'+
             ',0,QBI_NEWQTEVAL,0'+
             ',0,QBI_NEWCAVAL2,0'+
             ',0,QBI_NEWCAVAL3,0'+
             ',0,QBI_NEWCAVAL4,0'+
             ',0,QBI_NEWCAVAL5,0'+
             ',0,QBI_NEWCAVAL6,0';

 codeChpSqlW:='';
 for i:=1 to niv-1 do
  begin
   if BpOkOrli
    then codeChpSqlW:=codeChpSqlW+' AND QBI_VALAXENIV'+IntToStr(i)+'=QBQ_VALAXECT'+TabNumValAxe[i]
    else codeChpSqlW:=codeChpSqlW+' AND QBI_VALAXENIV'+IntToStr(i)+'=QBQ_VALAXECT'+IntToStr(i);
  end;
 if BpOkOrli
  then codeChpSqlW:=codeChpSqlW+' AND QBI_VALEURAXE=QBQ_VALAXECT'+TabNumValAxe[niv]
  else codeChpSqlW:=codeChpSqlW+' AND QBI_VALEURAXE=QBQ_VALAXECT'+IntToStr(niv);

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,QBQ_CODESESSION) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'"'+codeChpArbre+
            codeChpVal+',"'+codeSession+'"'+' FROM QBPSAISIEPREV '+
            ' WHERE QBI_CODESESSION="'+codeSession+
            '" AND QBI_NIVEAU="'+IntToStr(niv)+
            '" AND NOT EXISTS (SELECT QBQ_CODESESSION FROM QBPCUBETMP WHERE '+
             'QBQ_CODESESSION="'+codeSession+
             '" '+codeChpSqlW+')',
            'InsertCubeNewLigneArbre (RemplitTableQBPCubeTmpPgiGlobal).');
end;

procedure MAJCubeLigneSuppArbre(const codeSession:string;TabNumValAxe:array of string;
          niv:integer);
var i:integer;
    codeChpSql:string;
begin
 codeChpSql:='';
 for i:=1 to niv-1 do
  begin
   if BPOkOrli
    then codeChpSql:=codeChpSql+' AND QBI_VALAXENIV'+IntToStr(i)+'=QBQ_VALAXECT'+TabNumValAxe[i]
    else codeChpSql:=codeChpSql+' AND QBI_VALAXENIV'+IntToStr(i)+'=QBQ_VALAXECT'+IntToStr(i);
  end;
 if BPOkOrli
  then codeChpSql:=codeChpSql+' AND QBI_VALEURAXE=QBQ_VALAXECT'+TabNumValAxe[niv]
  else codeChpSql:=codeChpSql+' AND QBI_VALEURAXE=QBQ_VALAXECT'+IntToStr(niv);

 MExecuteSql('UPDATE QBPCUBETMP SET QBQ_PREVU=0,QBQ_CAPREVU=0,QBQ_CAPREVU2=0,'+
             'QBQ_CAPREVU3=0,QBQ_CAPREVU4=0,QBQ_CAPREVU5=0,QBQ_CAPREVU6=0 '+
             ' WHERE '+WhereCtx(['qbpcubetmp'])+
             ' AND QBQ_CODESESSION="'+codeSession+'" AND NOT EXISTS '+
             '(SELECT QBI_CODESESSION FROM QBPSAISIEPREV WHERE '+
             'QBI_CODESESSION="'+codeSession+'" AND QBI_NIVEAU="'+IntToStr(niv)+
             '" '+codeChpSql+')',
             'BPCalculCubeTmp (MAJCubeLigneSuppArbre).');
end;

procedure RemplitTableQBPCubeTmpPrevuObjGlobal(const codeSession:string;
          TabCodeAxe,TabNumValAxe:array of string);
var QTBP:TQTob;
    TobF:Tob;
    nivMax,i,j:integer;
    NewValeur,NewQte:double;
    codeC,codeP,joinP,ist,chpPP,joinPP,codeSql:string;
    Q:TQuery;
    CAVal1,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,QteVal:double;
    NewCAVal1,NewCAVal2,NewCAVal3,NewCAVal4,NewCAVal5,NewCAVal6,NewQteVal:double;
    RI:integer;
    codeSqlW:string;
begin
 QTBP:=TQTob.create(['AXE1','AXE2','AXE3','AXE4','AXE5',
                     'AXE6','AXE7','AXE8','AXE9','AXE10','NOEUD'],
                    ['VALCA1','NEWVALCA1','VAlCA2','NEWVALCA2',
                    'VALCA3','NEWVALCA3','VAlCA4','NEWVALCA4',
                    'VALCA5','NEWVALCA5','VALCA6','NEWVALCA6',
                    'VALQTE','NEWVALQTE']);
 //met l'arbre en ligne
 //on recupère que les lignes qui ont été modifiées
 //càd les lignes où histo<>newvaleur

 nivMax:=ChercheNivMax(codeSession);
 MetArbreEnLigne(nivMax,'VALIDATION',codeSession,0,TabCodeAxe,QTBP);

 RI:=DonneValeurAffiche(codeSession);

 //Mise à jour de la table qbpcubetmp
 //à partir de la tob contenant le prévu
 for i := 0 to QTBP.laTob.Detail.Count - 1 do
  begin
   TobF := QTBP.laTob.Detail[i];
   CAVal1:= TobF.GetValue('VALCA1');
   NewCAVal1 := TobF.GetValue('NEWVALCA1');
   CAVal2:= TobF.GetValue('VALCA2');
   NewCAVal2 := TobF.GetValue('NEWVALCA2');
   CAVal3:= TobF.GetValue('VALCA3');
   NewCAVal3 := TobF.GetValue('NEWVALCA3');
   CAVal4:= TobF.GetValue('VALCA4');
   NewCAVal4 := TobF.GetValue('NEWVALCA4');
   CAVal5:= TobF.GetValue('VALCA5');
   NewCAVal5 := TobF.GetValue('NEWVALCA5');
   CAVal6:= TobF.GetValue('VALCA6');
   NewCAVal6 := TobF.GetValue('NEWVALCA6');
   QteVal:= TobF.GetValue('VALQTE');
   NewQteVal := TobF.GetValue('NEWVALQTE');

   codeC:='';
   codeP:='';
   joinP:='';
   for j:=1 to nivMax do
    begin
     ist:=IntToStr(j);
     if BPOkOrli
      then codeC:=codeC+' AND QBQ_VALAXECT'+TabNumValAxe[j]+'="'+TobF.GetValue('AXE'+ist)+'" '
      else codeC:=codeC+' AND QBQ_VALAXECT'+ist+'="'+TobF.GetValue('AXE'+ist)+'" ';
     DonneCodeChpJoinPGI(TabCodeAxe[j],chpPP,joinPP);
     codeP:=codeP+' AND '+chpPP+'="'+TobF.GetValue('AXE'+ist)+'" ';
     joinP:=joinP+' '+joinPP;
    end;

   NewQte:=0;

   if not BPOkOrli
    then
     begin
      case RI of
       1 : begin
            NewValeur:=NewCAVal1-CAVal1;
            codeSqlW:='QBQ_CAPREVU';
           end;
       2 : begin
            NewValeur:=NewQteVal-QteVal;
            codeSqlW:='QBQ_PREVU';
           end;
       3 : begin
            NewValeur:=NewCAVal2-CAVal2;
            codeSqlW:='QBQ_CAPREVU2';
           end;
       4 : begin
            NewValeur:=NewCAVal3-CAVal3;
            codeSqlW:='QBQ_CAPREVU3';
           end;
       5 : begin
            NewValeur:=NewCAVal4-CAVal4;
            codeSqlW:='QBQ_CAPREVU4';
           end;
       6 : begin
            NewValeur:=NewCAVal5-CAVal5;
            codeSqlW:='QBQ_CAPREVU5';
           end;
       7 : begin
            NewValeur:=NewCAVal6-CAVal6;
            codeSqlW:='QBQ_CAPREVU6';
           end;
       else begin
             NewValeur:=NewCAVal1-CAVal1;
             codeSqlW:='QBQ_CAPREVU';
            end;
      end;
     end
    else
     begin
      NewValeur:=NewCAVal1-CAVal1;
      codeSqlW:='QBQ_CAPREVU';
      NewQte:=NewQteVal-QteVal;
     end;

   Q:=MOPenSQl('SELECT count(*) from qbpcubetmp where '+
               WhereCtx(['qbpcubetmp'])+' AND QBQ_CODESESSION="'+codeSession+
               '" AND '+codeC+' AND '+codeSqlW+'<>0 ','',true);
   if not Q.eof
    then
     begin
      if Q.fields[0].asFloat<>0
       then NewValeur:=NewValeur/(Q.fields[0].asFloat);
     end;
   ferme(Q);

   if BPPgiOrli
    then codeSql:=' '+codeSqlW+'='+codeSqlW+'+'+strFPoint4(NewValeur)
    else codeSql:=' QBQ_CAPREVU=QBQ_CAPREVU+'+strFPoint4(NewValeur)+
                  ',QBQ_PREVU=QBQ_PREVU+'+strFPoint4(NewQte);

   MExecuteSql('UPDATE QBPCUBETMP SET '+codeSql+
              ' WHERE '+WhereCtx(['qbpcubetmp'])+
              ' AND QBQ_CODESESSION="'+codeSession+'" AND'+
              codeC+' AND '+codeSqlW+'<>0',
             'BPCalculCubeTmp (RemplitTableQBPCubeTmpPrevuObjGlobal).');

  end;
 if QTBP <> nil
  then QTBP.Free;

 InsertCubeNewLigneArbre(codeSession,TabNumValAxe,nivMax);
 MAJCubeLigneSuppArbre(codeSession,TabNumValAxe,nivMax);

end;


//**********************************************************************
//
//                            Cas PGI
//
//**********************************************************************


procedure RemplitTableQBPCubeTmpPgiDelaiMaille(
          codeSession,codeChpSql,champs,ChpTotal,joinChpSql,
          codeRestriction:string;
          TabCodeAxe:array of string;
          DateDebCourante,DateFinCourante:TDateTime);
var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    ListMaille:TListMaille;
    i:integer;                                           
    Mi:TMaille;                                  
    types,codeChpValH,codeChpValR,whereH,whereR,table,join:string;
begin
 //cherche les dates de la session
 ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);

 ListMaille:=TListMaille.create();
 types:=SessionBPInitialise(codeSession);
 //if types=''
 // then
  types:='4';
 InitialiseListeMaille(StrToInt(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

 codeChpValR:=',SUM(GL_TOTALTTC),0,0,'+
              'SUM(GL_QTEFACT),0,0,'+
              'SUM(GL_TOTALHT),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0  ';
 codeChpValH:=',0,0,SUM(GL_TOTALTTC),'+
              '0,0,SUM(GL_QTEFACT),'+
              '0,0,SUM(GL_TOTALHT),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB))';
 Table:=' LIGNE ';
 join:='';
 WhereH:=' GL_DATEPIECE>="'+USDATETIME(DateDebRef)+
         '" AND GL_DATEPIECE<="'+USDATETIME(DateFinRef)+
         '" AND GL_TYPELIGNE="ART" ';
 WhereR:=' GL_DATEPIECE>="'+USDATETIME(DateDebCourante)+
         '" AND GL_DATEPIECE<="'+USDATETIME(DateFinCourante)+
         '" AND GL_TYPELIGNE="ART" ';
 if BPOkPgiEnt
  then
   begin
    codeChpValH:=',SUM(GL_TOTALTTC),0,0,'+
              'SUM(GSM_QPREVUE),0,0,'+
              'SUM(GL_TOTALHT),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0 ,'+
              'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0 ';
    codeChpValR:=',0,0,SUM(GL_TOTALTTC),'+
              '0,0,SUM(GSM_QPREVUE),'+
              '0,0,SUM(GL_TOTALHT),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB))';
    Table:=' STKMOUVEMENT ';
    WhereH:=' GSM_DATEPREVUE>="'+USDATETIME(DateDebRef)+
           '" AND GSM_DATEPREVUE<="'+USDATETIME(DateFinRef)+'" ';
    WhereR:=' GSM_DATEPREVUE>="'+USDATETIME(DateDebCourante)+
           '" AND GSM_DATEPREVUE<="'+USDATETIME(DateFinCourante)+'" ';
    join:='LEFT JOIN LIGNE ON GSM_NATUREORI=GL_NATUREPIECEG and '+
          'GSM_SOUCHEORI=GL_SOUCHE and GSM_NUMEROORI=GL_NUMERO and '+
          'GSM_INDICEORI=GL_INDICEG and GSM_NUMLIGNEORI=GL_NUMORDRE';
    if pos(join,joinChpSql)<>0
     then join:='';
   end;

 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);

   WhereH:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebReference)+
           '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinReference)+
           '" AND GL_TYPELIGNE="ART" ';
   WhereR:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebCourante)+
           '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinCourante)+
           '" AND GL_TYPELIGNE="ART" ';
   if BPOkPgiEnt
    then
     begin   
      WhereH:=' GSM_DATEPREVUE>="'+USDATETIME(Mi.DateDebReference)+
              '" AND GSM_DATEPREVUE<="'+USDATETIME(Mi.DateFinReference)+'" ';
      WhereR:=' GSM_DATEPREVUE>="'+USDATETIME(Mi.DateDebCourante)+
              '" AND GSM_DATEPREVUE<="'+USDATETIME(Mi.DateFinCourante)+'" ';
     end;
   //*************** insertion réalisé
   //on prend dans la table ligne
   //pour les dates de la maille et code restriction
   MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
            codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
            '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
            '" '+codeChpValR+
            ' FROM '+Table+joinChpSql+' '+join+
            ' WHERE '+whereR+codeRestriction+
            ' GROUP BY '+champs+',GL_DATEPIECE',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPgiDelaiMaille).');

    //*************** insertion histo
    //on prend dans la table ligne
    MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'","'+
            STRFPOINT(Mi.code)+'",'+champs+
            ',"'+USDATETIME(Mi.DateDebCourante)+'"'+CodeChpValH+
            ' FROM '+Table+joinChpSql+' '+join+
            ' WHERE '+WhereH+codeRestriction+
            ' GROUP BY '+champs+',GL_DATEPIECE',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPgiDelaiMaille).');

     //**************** insertion prévu
     RemplitTableQBPCubeTmpPrevuObjDelaiMaille(codeSession,codeChpSql,
                                     Mi,TabCodeAxe);

  end;
end;


procedure RemplitTableQBPCubeTmpPgiGlobal(
          codeSession,codeChpSql,champs,ChpTotal,joinChpSql,
          codeRestriction:string;
          TabCodeAxe,TabNumValAxe:array of string;
          DateDebCourante,DateFinCourante:TDateTime);
var codeChpValR,codeChpValH,Table,join,WhereH,WhereR,ChpDate:string;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
begin
 ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);
 codeChpValR:=',SUM(GL_TOTALTTC),0,0,'+
            'SUM(GL_QTEFACT),0,0,'+
            'SUM(GL_TOTALHT),0,0,'+
            'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),0,0,'+
            'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),0,0,'+
            'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0,'+
            'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0 ';
 codeChpValH:=',0,SUM(GL_TOTALTTC),SUM(GL_TOTALTTC),'+
            '0,SUM(GL_QTEFACT),SUM(GL_QTEFACT),'+
            '0,SUM(GL_TOTALHT),SUM(GL_TOTALHT),'+
            '0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
            '0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
            '0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
            '0,SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)) ';
 Table:=' LIGNE ';
 join:='';
 WhereH:=' GL_DATEPIECE>="'+USDATETIME(DateDebRef)+
         '" AND GL_DATEPIECE<="'+USDATETIME(DateFinRef)+'" ';
 WhereR:=' GL_DATEPIECE>="'+USDATETIME(DateDebCourante)+
         '" AND GL_DATEPIECE<="'+USDATETIME(DateFinCourante)+'" ';
 ChpDate:=' ,GL_DATEPIECE ';
 if BPOkPgiEnt
  then
   begin
    codeChpValH:=',0,SUM(GL_TOTALTTC),SUM(GL_TOTALTTC),'+
              '0,SUM(GSM_QPREVUE),SUM(GSM_QPREVUE),'+
              '0,SUM(GL_TOTALHT),SUM(GL_TOTALHT),'+
              '0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
              '0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
              '0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
              '0,SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)) ';
    codeChpValR:=',SUM(GL_TOTALTTC),0,0,'+
              'SUM(GSM_QPREVUE),0,0,'+
              'SUM(GL_TOTALHT),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0 ';
    Table:=' STKMOUVEMENT ';
    WhereH:=' GSM_DATEPREVUE>="'+USDATETIME(DateDebRef)+
           '" AND GSM_DATEPREVUE<="'+USDATETIME(DateFinRef)+'" ';
    WhereR:=' GSM_DATEPREVUE>="'+USDATETIME(DateDebCourante)+
           '" AND GSM_DATEPREVUE<="'+USDATETIME(DateFinCourante)+'" ';
    join:='LEFT JOIN LIGNE ON GSM_NATUREORI=GL_NATUREPIECEG and '+
          'GSM_SOUCHEORI=GL_SOUCHE and GSM_NUMEROORI=GL_NUMERO and '+
          'GSM_INDICEORI=GL_INDICEG and GSM_NUMLIGNEORI=GL_NUMORDRE';
    if pos(join,joinChpSql)<>0
     then join:='';
    ChpDate:=' ,GSM_DATEPREVUE ';
   end;


 //*************** insertion réalisé
 //on prend dans la table ligne
 //pour les dates courantes et code restriction
 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,QBQ_CODESESSION) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'",'+champs+ChpDate+
            codeChpValR+',"'+codeSession+'"'+
            ' FROM '+Table+joinChpSql+' '+join+
            ' WHERE '+whereR+codeRestriction+
            ' GROUP BY '+champs+ChpDate,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPgiGlobal).');


 //*************** insertion histo
 //on prend dans la table ligne
 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,QBQ_CODESESSION) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'",'+champs+ChpDate+'+365'+
            codeChpValH+',"'+codeSession+'"'+' FROM '+table+joinChpSql+' '+join+
            ' WHERE '+whereH+codeRestriction+
            ' GROUP BY '+champs+ChpDate,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPgiGlobal).');

 //*************** insertion prévu
 RemplitTableQBPCubeTmpPrevuObjGlobal(codeSession,TabCodeAxe,TabNumValAxe);
end;

//remplit table qbpcubetmp
//dans le cas pgi
procedure RemplitTableQBPCubeTmpPgi(const codeSession:string;
          DateDebCourante,DateFinCourante:TDateTime);
var codeChpSql,ChpTotal,codeRestriction,champs,joinChpSql:string;
    chpLib,joinLib,codeJoinR:string;
    nivMax,nivMax0,i:integer;
    TabCodeAxe,TabNumValAxe:array [1..11] of string;
begin
 //vide la table temporaire
 MExecuteSql('DELETE FROM QBPCUBETMP WHERE '+WhereCtx(['qbpcubetmp'])+
            ' AND QBQ_CODESESSION="'+codeSession+'"',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpPgi).');

 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 if not SessionInitPrev(codesession)
  then nivMax:=ChercheNivMax(codeSession)
  else
   begin
    nivMax0:=ChercheNivMax(codeSession); 
    nivMax0:=nivMax0+2;
    TabCodeAxe[NivMax0]:='ARTICLE';
    if GestionBPColoris
     then
      begin
       nivMax0:=nivMax0+1;
       TabCodeAxe[nivMax0]:='COLORIS';
      end;
    //niv FS
    if GestionBPFS
     then 
      begin
       nivMax0:=nivMax0+1;
       TabCodeAxe[nivMax0]:='FS';
      end;
    //niv Magasin
    if GestionBPMagasin
     then 
      begin
       nivMax0:=nivMax0+1;
       TabCodeAxe[nivMax0]:='MAGASIN';
       if BPOkPgiEnt
        then TabCodeAxe[nivMax0]:='DEPOT';
      end;
    nivMax:=ChercheNivMaxSession(codeSession);
   end;

 //cherche les codes pour le sql
 codeChpSql:='';
 for i:=1 to nivMax do
  begin
   if codeChpSql=''
    then codeChpSql:='QBQ_VALAXECT'+IntToStr(i)
    else codeChpSql:=codeChpSql+',QBQ_VALAXECT'+IntToStr(i);
  end;
 for i:=1 to nivMax do
  begin
   codeChpSql:=codeChpSql+',QBQ_LIBVALAXECT'+IntToStr(i);
  end;

 DonneCodeChpJoinPGINbNiv(TabCodeAxe,nivMax,champs,joinChpSql);
 ChercheCodeSqlChpAxeLibPGI(nivMax,TabCodeAxe,chpLib,joinLib);
 codeRestriction:=ChercheCodeRestrictionSession('','','',codeSession,codeJoinR);

 //remplit table tmp cube
 //avec l'histo, le realisé et le prévu
 if SessionDelai(codesession)
  then
   begin
    //session eclatée
    RemplitTableQBPCubeTmpPgiDelaiMaille(codeSession,codeChpSql,champs+chpLib,ChpTotal,
           joinChpSql+' '+joinLib,codeRestriction,TabCodeAxe,
           DateDebCourante,DateFinCourante);
   end
  else
   begin
    //session globale
    RemplitTableQBPCubeTmpPgiGlobal(codeSession,codeChpSql,champs+chpLib,ChpTotal,
           joinChpSql+' '+joinLib,codeRestriction,
           TabCodeAxe,TabNumValAxe,
           DateDebCourante,DateFinCourante);
   end;
end ;


//**********************************************************************
//
//                            Cas ORLI
//
//**********************************************************************


procedure RemplitTableQBPCubeTmpOrliHistoArbre();
begin

end;

//procedure qui met dans la table Tmp Cube
//l'historique par rapport à la table pivot
procedure RemplitTableQBPCubeTmpOrliHistoPivot(const codeSession,saiCmd,saiCmdRef,codeEclat,codeChpSql,codeP,codeB,
          structure,codeRestrictionP,codeRestrictionB:string;
          DateDebRef,DateFinRef:TDateTime);
var WhereH:string;
begin
 WhereH:=' QBV_VALAXEP8="'+saiCmdRef+'" ';
 if SaiCmdRef=''
  then WhereH:=' QBV_DATEPIVOT>="'+USDATETIME(DateDebRef)+
               '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinRef)+'" ';
 if codeEclat='true'
  then MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+',QBQ_DATECT,'+
           'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
           'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+codeP+
           ',QBV_DATEPIVOT+365,SUM(QBV_QTE2),0,0,'+
           'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ',0,0 FROM QBPPIVOT '+
           'WHERE QBV_CODESTRUCT="'+structure+'" AND '+WhereH+codeRestrictionP+
           ' GROUP BY '+codeP+',QBV_DATEPIVOT',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrliHistoPivot).')
  else MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+',QBQ_DATECT,'+
           'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
           'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+codeP+
           ',QBV_DATEPIVOT+365,SUM(QBV_QTE2),0,SUM(QBV_QTE2),'+
           'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ',0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ' FROM QBPPIVOT '+
           'WHERE QBV_CODESTRUCT="'+structure+'" AND'+WhereH+codeRestrictionP+
           ' GROUP BY '+codeP+',QBV_DATEPIVOT',
           'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrliHistoPivot).');
end;



procedure RemplitTableQBPCubeTmpCas1(const codeSession,structure,codeChpSql,codeP,
          codeRestrictionP,chpSqlLib,joinSqlLib:string;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
          TabCodeAxe,TabNumValAxe:array of string;
          SaiCmd,SaiRef:string);
var code:string;
begin
 //*************** insertion réalisé
 //on prend dans la table qbppivot
 //pour la structure et les dates courantes (ou saison de commande)
 if SaiCmd<>''
  then code:=' QBV_VALAXEP8="'+SaiCmd+'" '
  else code:=' QBV_DATEPIVOT>="'+USDATETIME(DateDebC)+
             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinC)+'" ';

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+','+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+codeP+chpSqlLib+','+
            'SUM(QBV_QTE2),0,0,'+
            'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
            'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
            ',0,0 FROM QBPPIVOT '+joinSqlLib+
            ' WHERE QBV_CODESTRUCT="'+structure+'" AND '+code+codeRestrictionP+
            ' GROUP BY '+codeP+chpSqlLib,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas1).');

 //*************** insertion histo
 if SaiRef<>''
  then code:=' QBV_VALAXEP8="'+SaiRef+'" '
  else code:=' QBV_DATEPIVOT>="'+USDATETIME(DateDebRef)+
             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinRef)+'" ';

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+','+
           'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
           'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+codeP+chpSqlLib+
           ',SUM(QBV_QTE2),0,SUM(QBV_QTE2),'+
           'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ',0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ' FROM QBPPIVOT '+joinSqlLib+
           ' WHERE QBV_CODESTRUCT="'+structure+'" AND '+code+codeRestrictionP+
           ' GROUP BY '+codeP+chpSqlLib,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas1).');

 //*************** insertion prévu
 RemplitTableQBPCubeTmpPrevuObjGlobal(codeSession,TabCodeAxe,TabNumValAxe);

end;

procedure RemplitTableQBPCubeTmpCas2(const codeSession,structure,codeChpSql,codeP,
          codeRestrictionP,chpSqlLib,joinSqlLib:string;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
          TabCodeAxe,TabNumValAxe:array of string);
var ListMaille:TListMaille;
    i:integer;
    Mi:TMaille;
    types:string;
begin
 ListMaille:=TListMaille.create();
 SessionCalculParDelai(codeSession,types);
 InitialiseListeMaille(StrToInt(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

 //pour chaque maille
 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);

   //*************** insertion réalisé
   //on prend dans la table qbppivot
   //pour la structure et les dates courantes
   MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+
              ',QBQ_DATECT,'+
              'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
              'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+chpSqlLib+',"'+
              USDATETIME(Mi.DateDebCourante)+'",'+
              'SUM(QBV_QTE2),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
              USDATETIME(Mi.DateDebCourante)+
              '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinCourante)+
              '" '+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas2).');

   //*************** insertion histo
   MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+
              ',QBQ_DATECT,'+
              'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
              'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+chpSqlLib+
              ',"'+USDATETIME(Mi.DateDebCourante)+'",SUM(QBV_QTE2),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
              USDATETIME(Mi.DateDebReference)+
              '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinReference)+
              '" '+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas2).');
  end;

 //*************** insertion prévu
 RemplitTableQBPCubeTmpPrevuObjDelai(codeSession,codeChpSql,'',
                                     TabCodeAxe);
end;

procedure RemplitTableQBPCubeTmpCas3();
begin

end;

procedure RemplitTableQBPCubeTmpCas4(const codeSession,structure,codeChpSql,codeP,
          codeRestrictionP,chpSqlLib,joinSqlLib:string;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
          TabCodeAxe,TabNumValAxe:array of string);
var SessionObj,codesqlArbre:string;
    nivMax,i:integer;
begin
 //*************** insertion réalisé
 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+','+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
            '",'+codeP+','+
            'SUM(QBV_QTE2),0,0,'+
            'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
            'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
            ',0,0 FROM QBPPIVOT '+
            'WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
            USDATETIME(DateDebC)+
            '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinC)+
            '" '+codeRestrictionP+
            ' GROUP BY '+codeP+',QBV_DATEPIVOT',
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas4).');

 //*************** insertion histo
 if ChercheSessionObjectif(codesession,SessionObj)
  then
   begin
    //si la prevision est faite à partir d'un objectif
    //on lit l'histo dans table arbre objectif
    nivMax:=ChercheNivMax(codeSession);

    codesqlArbre:='';

    for i:=1 to NivMax-1 do
     begin
      if codesqlArbre=''
       then codesqlArbre:=' QBI_VALAXENIV'+IntToStr(i)
       else codesqlArbre:=codesqlArbre+',QBI_VALAXENIV'+IntToStr(i);
     end;

    MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+','+
               'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
               'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
               'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
               'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
               'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
               'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5) '+
               ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
               codesqlArbre+','+
               '0,QBI_NEWVAL,0,'+
               '0,QBI_NEWQTEVAL,0,'+
               '0,QBI_NEWCAVAL2,0,'+
               '0,QBI_NEWCAVAL3,0,'+
               '0,QBI_NEWCAVAL4,0,'+
               '0,QBI_NEWCAVAL5,0 '+
               ' FROM QBPSAISIEPREV '+
               ' WHERE QBI_CODESESSION="'+codeSession+
               '" AND QBI_NIVEAU="'+IntToStr(NivMax)+'" ',
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas4).');
   end
  else
   begin
    //prevision faite à partir de l'histo de la table qbppivot
    MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+','+
           'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
           'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+codeP+chpSqlLib+
           ',SUM(QBV_QTE2),0,SUM(QBV_QTE2),'+
           'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ',0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
           'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
           'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
           'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
           ' FROM QBPPIVOT '+joinSqlLib+
           ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
           USDATETIME(DateDebRef)+
           '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinRef)+
           '" '+codeRestrictionP+
           ' GROUP BY '+codeP+chpSqlLib,
           'BPCalculCubeTmp (RemplitTableQBPCubeTmpCas4).');
   end;
 
 //*************** insertion prévu


end;

procedure RemplitTableQBPCubeTmpCas5();
begin

end;

procedure RemplitTableQBPCubeTmpCas6();
begin

end;

procedure RemplitTableQBPCubeTmpOrli0(const codeSession,codeChpSql,codeP,codeB,
          codeChpSqlS,codeChpSqlLib,codePS,codeBS,
          structure,codeRestrictionP,codeRestrictionB,chpSqlLib,joinSqlLib:string;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
          TabCodeAxe,TabNumValAxe:array of string;
          SaiCmd,SaiRef:string);
var SessionObj:string;
    codesqlArbre:string;
    nivMax,i:integer;
    ListMaille:TListMaille;
    Mi:TMaille;
    types,chpArbreLib,joinArbreLib,codefiltre:string;
begin
 ListMaille:=TListMaille.create();
 SessionCalculParDelai(codeSession,types);
 InitialiseListeMaille(StrToInt(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

 //pour chaque maille
 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);

   //*************** insertion réalisé
   //on prend dans la table qbppivot
   //pour la structure et les dates courantes
   MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+
              ',QBQ_DATECT,'+
              'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
              'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+',"'+
              USDATETIME(Mi.DateDebCourante)+'",'+
              'SUM(QBV_QTE2),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
              USDATETIME(Mi.DateDebCourante)+
              '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinCourante)+
              '" '+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrli0).');
   //*************** insertion histo

   //si la session est une prevision
   //et qu'elle n'est pas initialisée par rapport à un objectif
   if not ChercheSessionObjectif(codesession,SessionObj)
    then MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+
              ',QBQ_DATECT,'+
              'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
              'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+
              ',"'+USDATETIME(Mi.DateDebCourante)+'",SUM(QBV_QTE2),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
              USDATETIME(Mi.DateDebReference)+
              '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinReference)+
              '" '+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrli0).');
  end;

 //*************** insertion prévu
 nivMax:=ChercheNivMax(codeSession);
 if SessionDelai(codeSession)
  then nivMax:=nivMax+1;
 codesqlArbre:='';

 for i:=1 to NivMax-1 do
  begin
   if codesqlArbre=''
    then codesqlArbre:=' QBR_VALAXENIV'+IntToStr(i)
    else codesqlArbre:=codesqlArbre+',QBR_VALAXENIV'+IntToStr(i);
  end;

 ChercheCodeSqlChpAxeLib(nivMax,TabCodeAxe,chpArbreLib,joinArbreLib);

 MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
            'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
            codesqlArbre+chpArbreLib+',QBR_DATEDELAI,'+
            '0,QBR_OP1,0,'+
            '0,QBR_QTEC,0,'+
            '0,QBR_OP2,0,'+
            '0,QBR_OP3,0,'+
            '0,QBR_OP4,0,'+
            '0,QBR_OP5,0,'+
            '0,QBR_OP6,0 '+
            ' FROM QBPARBRE '+joinArbreLib+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NIVEAU="'+IntToStr(NivMax)+'" '+codefiltre,
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrli0).'); 
end;


//remplit table qbpcubetmp
//dans le cas orli
procedure RemplitTableQBPCubeTmpOrli(const codeSession:string;
          DateDebCourante,DateFinCourante:TDateTime);
var codeRestrictionP,codeRestrictionB:string;
    codeChpSql,codeP,codeB,structure:string;
    codeChpSqlS,codeChpSqlLib,codePS,codeBS,chpSqlLib,joinSqlLib:string;
    i:integer;
    nivMax:integer;
    TabCodeAxe,TabNumValAxe:array [1..11] of string;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    SaiCmd,SaiRef,codeJoinR:string;
begin
 if (not OkSessionModif(codeSession)) and (ExisteSql('SELECT * FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+codeSession+'"'))
  then exit;

 //vide la table temporaire
 MExecuteSql('DELETE FROM QBPCUBETMP WHERE '+WhereCtx(['qbpcubetmp'])+
            ' AND QBQ_CODESESSION="'+codeSession+'"',
            'BPCalculCubeTmp (RemplitTableQBPCubeTmpOrli).');

 structure:=ChercheSessionStructure(codeSession);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche les codes pour le sql
 codeP:='';
 codeChpSql:='';
 codeB:='';
 codePS:='';
 codeChpSqlS:='';
 codeBS:='';

 //cube avec axe session
 nivMax:=ChercheNivMax(codeSession);
 for i:=1 to nivMax do
  begin
   if codeBS=''
    then
     begin
      codeBS:='QBB_VALAXEBP'+TabNumValAxe[i+1];
      codePS:='QBV_VALAXEP'+TabNumValAxe[i+1];
      codeChpSqlS:='QBQ_VALAXECT'+TabNumValAxe[i+1];
     end
    else
     begin
      codeBS:=codeBS+',QBB_VALAXEBP'+TabNumValAxe[i+1];
      codePS:=codePS+',QBV_VALAXEP'+TabNumValAxe[i+1];
      codeChpSqlS:=codeChpSqlS+',QBQ_VALAXECT'+TabNumValAxe[i+1];
     end;
  end;
 //libellé
 codeChpSqlLib:=codeChpSqlS;
 for i:=1 to nivMax do
  codeChpSqlLib:=codeChpSqlLib+',QBQ_LIBVALAXECT'+TabNumValAxe[i+1];

  for i:=1 to 30 do
  begin
   if codeB=''
    then
     begin
      codeB:='QBB_VALAXEBP'+IntToStr(i);
      codeP:='QBV_VALAXEP'+IntToStr(i);
      codeChpSql:='QBQ_VALAXECT'+IntToStr(i);
     end
    else
     begin
      codeB:=codeB+',QBB_VALAXEBP'+IntToStr(i);
      codeP:=codeP+',QBV_VALAXEP'+IntToStr(i);
      codeChpSql:=codeChpSql+',QBQ_VALAXECT'+IntToStr(i);
     end;
  end; 

 //libellé
{ for i:=1 to 30 do
  codeChpSql:=codeChpSql+',QBQ_LIBVALAXECT'+IntToStr(i);
}
 codeRestrictionB:=ChercheCodeRestrictionSession('','QBB_VALAXEBP','QBB_',codeSession,codeJoinR);
 codeRestrictionP:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,codeJoinR);

 //cherche dates periode courante et reference de la session
 ChercheDateDDateFPeriode(codeSession,
                          DateDebC,DateFinC,DateDebRef,DateFinRef);
 //cherche saison de commande et de reference
 ChercheSaisonCmd(codeSession,SaiCmd,SaiRef);

 //code pour avoir les libelles
 ChercheCodeSqlChpAxeLibAvecPivot(NivMax,'',TabCodeAxe,TabNumValAxe,
                                  chpSqlLib,joinSqlLib);

 //remplit table tmp cube
 //avec l'histo, le realisé et le prévu
 RemplitTableQBPCubeTmpOrli0(codeSession,codeChpSql,codeP,codeB,
          codeChpSqlS,codeChpSqlLib,codePS,codeBS,
          structure,codeRestrictionP,codeRestrictionB,chpSqlLib,joinSqlLib,
          DateDebC,DateFinC,DateDebRef,DateFinRef,TabCodeAxe,TabNumValAxe,
          SaiCmd,SaiRef);

 ModifSession(codeSession,'-');

end;

end.
