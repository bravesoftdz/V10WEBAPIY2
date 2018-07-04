unit BPCoeff;

interface

uses HEnt1;

procedure CalculCoeff(const codeSession:hstring);

implementation

uses Sysutils,HCtrls,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     UTob,UGraph,CstCommun,Uutil,BPBasic,BPFctSession,BPVueClient;

//*****************************************************************************
//calcul coeff saisonnier

procedure MiseAJourDesCoeffsAvecCoeffGene(const codeSession:hstring);
var Q:TQuery;
    SeuilExtrapo,SommeCaTotal,SommeCAExtrapo,CoeffGene:double;
begin
 SeuilExtrapo:=DonneParamR(pr_BPSeuilExtrapo);
 //Somme CA extrapolable
 SommeCAExtrapo:=1;
 Q:=MOpenSql('SELECT SUM(QBT_CARETENU) FROM QBPCOEFF '+
             'WHERE QBT_CODESESSION="'+codeSession+
             '" AND QBT_COEFFEXTRAPO<='+STRFPOINT(seuilExtrapo)+
             ' AND QBT_COEFFEXTRAPO<>0 AND QBT_EXTRAPOLABLE="1"','BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).',true);
 if not Q.eof
  then SommeCAExtrapo:=Q.fields[0].asFloat;
 ferme(Q);
 //Somme CA total
 SommeCaTotal:=0;
 Q:=MOpenSql('SELECT SUM(QBT_CARETENU) FROM QBPCOEFF '+
             'WHERE QBT_CODESESSION="'+codeSession+
             '" AND QBT_EXTRAPOLABLE<>"2" ','BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).',true);
 if not Q.eof
  then SommeCaTotal:=Q.fields[0].asFloat;
 ferme(Q);

 Q:=MOpenSql('SELECT SUM(QBT_CAREELVU+QBT_CAREELNEW) FROM QBPCOEFF '+
             'WHERE QBT_CODESESSION="'+codeSession+
             '" AND QBT_COEFFEXTRAPO>'+STRFPOINT(seuilExtrapo)+
             ' AND QBT_EXTRAPOLABLE<>"2"','BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).',true);
 if not Q.eof
  then SommeCaTotal:=SommeCaTotal-Q.fields[0].asFloat;
 ferme(Q);
 //coeff gene
 if SommeCAExtrapo<>0
  then CoeffGene:=SommeCaTotal/SommeCAExtrapo
  else CoeffGene:=1;
 //mise à jour des coeffs
 MExecuteSql('UPDATE QBPCOEFF SET '+
            'QBT_COEFFRETENU=QBT_COEFFRETENU*'+STRFPOINT(CoeffGene)+
            ' WHERE QBT_CODESESSION="'+codeSession+
            '" AND QBT_COEFFEXTRAPO<='+STRFPOINT(seuilExtrapo),
            'BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).');
 MExecuteSql('UPDATE QBPCOEFF SET '+
            'QBT_COEFFRETENU=1 '+
            ' WHERE QBT_CODESESSION="'+codeSession+
            '" AND QBT_COEFFEXTRAPO>'+STRFPOINT(seuilExtrapo),
            'BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).');
 MExecuteSql('UPDATE QBPCOEFF SET '+
            'QBT_COEFFRETENU=1 '+
            ' WHERE QBT_CODESESSION="'+codeSession+
            '" AND QBT_EXTRAPOLABLE="2"',
            'BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).');
 MExecuteSql('UPDATE QBPCOEFF SET '+
            'QBT_COEFFGENE='+STRFPOINT(CoeffGene)+
            ' WHERE QBT_CODESESSION="'+codeSession+
            '" ',
            'BPCoeff (MiseAJourDesCoeffsAvecCoeffGene).');

end;

procedure MAJExtrapolable(codesession:hstring);
var Q:TQuery;
    TabCodeAxe:array [1..11] of hstring;
    TabNumValAxe:array [1..11] of hstring;
    nivMax,j,i:integer;
    code,ist,codeS,codeU:hstring;
begin
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);


 code:='';
 codeS:='';
 codeU:='';
 for j:=1 to NivMax+1 do
  begin
   ist:=IntToStr(j);
   if TabNumValAxe[j]<>''
    then
     begin
     code:=code+' AND QBV_VALAXEP'+TabNumValAxe[j]+'=QBT_CODEAXEC'+IntToStr(j-1);
     codeS:=codeS+ ',QBV_VALAXEP'+TabNumValAxe[j];
     end;
  end;
 Q:=MOpenSql('SELECT DISTINCT QBV_EXTRAPOLABLE'+codeS+
             ' FROM QBPPIVOT,QBPCOEFF WHERE '+code,
             'BPCoeff (MAJExtrapolable).',true );
 while not Q.eof do
  begin
   codeu:='';
   for i:=1 to NivMax do
   begin
    codeu:=codeU+ ' AND QBT_CODEAXEC'+intToStr(i)+'="'+Q.fields[i].asString+'"';
   end;

   MExecuteSql('UPDATE QBPCOEFF SET QBT_EXTRAPOLABLE="'+Q.fields[0].asString+
               '" WHERE '+codeU,
               'BPCoeff (MAJExtrapolable).');
   Q.next;
  end;
 ferme(Q);
end;

procedure RemplitTableQBP(const codeSession:hstring;
          QTBP:TQTob);
var i:integer;
    TobF:Tob;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:hstring;
    lib1,lib2,lib3,lib4,lib5,lib6,lib7,lib8,lib9,lib10,extra:hstring;
    cahistovu,cahistononvu,careelvu,careelnonvu,coeff,cahisto,CAProspect:double;
    Xvisit,NewProspect,ResteAVisiter,caRetenu:double;
    nbCltHisto,nbCltVu,nbCltHistoNonVu,nbCltHistoReelNonVu:integer;
    PrctSeuilCaHVu,PrctCaHVu,careelvuX:double;
    SessionObjectif:hstring;
begin
 PrctSeuilCaHVu:=DonneParamR(pr_BPPrctSeuilCaHVu);
 for i := 0 to QTBP.laTob.Detail.Count - 1 do
  begin
   TobF := QTBP.laTob.Detail[i];
   axe1 := TobF.GetValue('AXE1');
   axe2 := TobF.GetValue('AXE2');
   axe3 := TobF.GetValue('AXE3');
   axe4 := TobF.GetValue('AXE4');
   axe5 := TobF.GetValue('AXE5');
   axe6 := TobF.GetValue('AXE6');
   axe7 := TobF.GetValue('AXE7');
   axe8 := TobF.GetValue('AXE8');
   axe9 := TobF.GetValue('AXE9');
   axe10 := TobF.GetValue('AXE10');
   lib1 := TobF.GetValue('LIB1');
   lib2 := TobF.GetValue('LIB2');
   lib3 := TobF.GetValue('LIB3');
   lib4 := TobF.GetValue('LIB4');
   lib5 := TobF.GetValue('LIB5');
   lib6 := TobF.GetValue('LIB6');
   lib7 := TobF.GetValue('LIB7');
   lib8 := TobF.GetValue('LIB8');
   lib9 := TobF.GetValue('LIB9');
   lib10 := TobF.GetValue('LIB10'); 

   lib1 := '';
   lib2 := '';
   lib3 := '';
   lib4 := '';
   lib5 := '';
   lib6 := '';
   lib7 := '';
   lib8 := '';
   lib9 := '';
   lib10 := '';

   cahistovu := TobF.GetValue('CAHISTOVU');
   cahistononvu := TobF.GetValue('CAHISTONONVU');
   careelvu := TobF.GetValue('CAREELVU');
   careelnonvu := TobF.GetValue('CAREELNONVU');
   if BPOkOrli
    then
     //-----------------> ORLI
     begin
      if not ChercheSessionObjectif(codeSession,SessionObjectif)
       then cahisto := TobF.GetValue('CAHISTO')+cahistovu //cahistononvu+
       else cahisto := TobF.GetValue('CAHISTO')
     end
     //ORLI <-----------------
    else cahisto := TobF.GetValue('CAHISTO')+cahistononvu+cahistovu;
   nbCltVu := TobF.GetValue('NBCLIENTVU');
   nbCltHistoNonVu := TobF.GetValue('NBCLIENTHISTONONVU');
   nbCltHisto := TobF.GetValue('NBCLIENTHISTO')+nbCltVu+nbCltHistoNonVu;
   nbCltHistoReelNonVu := TobF.GetValue('NBCLIENTREELNONVU');
   CAProspect := TobF.GetValue('CAPROSPECT');      
   extra := FloatToStr(TobF.GetValue('EXTRA'));
   careelvuX:=careelvu-careelnonvu;
   if careelvuX<0
    then careelvuX:=0;
   if cahistovu<>0
    then Xvisit:=careelvuX/cahistovu
    else Xvisit:=1;
   if CAProspect>careelnonvu
    then NewProspect:=CAProspect
    else NewProspect:=0;//careelnonvu;
   ResteAVisiter:=cahisto-CAProspect-cahistovu-cahistononvu;
   //compare le % du ca histo vu par rapport au (ca histo - ca prospect)
   //avec le % seuil definit en paramètre
   //si < alors XVisit=1
   PrctCaHVu:=100;
   if cahisto-caprospect<>0
    then PrctCaHVu:=(cahistovu/(cahisto-caprospect)*100);
   if PrctCaHVu<PrctSeuilCaHVu
    then Xvisit:=1;

   caretenu:=(ResteAVisiter*Xvisit)+careelvu+NewProspect;

//   if coeff<>1
//    then
//     begin
      //on prend que les lignes qui sont extrapolables
      if (careelvu+careelnonvu)<>0
       then coeff:=caretenu/(careelvu+careelnonvu)
       else coeff:=0;

      MExecuteSql('INSERT INTO QBPCOEFF (QBT_CODESESSION,QBT_NUMCOEFF,'+
               'QBT_CODEAXEC1,QBT_CODEAXEC2,QBT_CODEAXEC3,QBT_CODEAXEC4,'+
               'QBT_CODEAXEC5,QBT_CODEAXEC6,QBT_CODEAXEC7,QBT_CODEAXEC8,'+
               'QBT_CODEAXEC9,QBT_CODEAXEC10,'+
               'QBT_LIBC1,QBT_LIBC2,QBT_LIBC3,QBT_LIBC4,'+
               'QBT_LIBC5,QBT_LIBC6,QBT_LIBC7,QBT_LIBC8,'+
               'QBT_LIBC9,QBT_LIBC10,QBT_CAHISTOVU,QBT_CAHISTONONVU,'+
               'QBT_CAREELVU,QBT_CAREELNEW,'+
               'QBT_COEFFEXTRAPO,QBT_CARETENU,QBT_COEFFRETENU,QBT_CAHISTO,'+
               'QBT_NBCLTHISTO,QBT_NBCLTVU,QBT_NBCLTHISTONVU,'+
               'QBT_NBCLTREELNONVU,QBT_CAPROSPECT,QBT_XVISIT,'+
               'QBT_NEWPROSPECT,QBT_RESTEAVISITER,QBT_EXTRAPOLABLE) ' +
               'VALUES ("'+codeSession+'","'+strFPoint4(i)+'","'+axe1+'","'+axe2+
               '","'+axe3+'","'+axe4+'","'+axe5+'","'+axe6+'","'+
               axe7+'","'+axe8+'","'+axe9+'","'+axe10+'","'+
               lib1+'","'+lib2+
               '","'+lib3+'","'+lib4+'","'+lib5+'","'+lib6+'","'+
               lib7+'","'+lib8+'","'+lib9+'","'+lib10+'","'+
               strFPoint4(cahistovu)+
               '","'+strFPoint4(cahistononvu)+'","'+strFPoint4(careelvu)+
               '","'+strFPoint4(careelnonvu)+
               '","'+strFPoint4(coeff)+'","'+strFPoint4(caretenu)+
               '","'+strFPoint4(coeff)+'","'+strFPoint4(cahisto)+
               '","'+strFPoint4(nbCltHisto)+'","'+strFPoint4(nbCltVu)+
               '","'+strFPoint4(nbCltHistoNonVu)+
               '","'+strFPoint4(nbCltHistoReelNonVu)+
               '","'+strFPoint4(CAProspect)+
               '","'+strFPoint4(Xvisit)+
               '","'+strFPoint4(NewProspect)+
               '","'+strFPoint4(ResteAVisiter)+'","'+extra+'")',
               'BPCoeff (RemplitTableQBP).');
    // end;
  end;
end;

//

procedure ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe:array of hstring;
          nivMax:integer;
          var codeChp,codeJoin,codeE:hstring);
var j:integer;
    code,chpSqlLib:hstring;
begin
 code:='';
 codeE:='';
 for j:=1 to 10 do
  begin
   if TabNumValAxe[j]<>''
    then
     begin
      if code<>''
       then code:=code+' ,R.QBV_VALAXEP'+TabNumValAxe[j]
       else code:=' R.QBV_VALAXEP'+TabNumValAxe[j];
      codeE:= codeE+' AND R.QBV_VALAXEP'+TabNumValAxe[j]+
              '=C.QBV_VALAXEP'+TabNumValAxe[j]+' ';
     end;
  end;

 //code pour avoir les libelles
 ChercheCodeSqlChpAxeLibAvecPivot(NivMax,'R.',TabCodeAxe,TabNumValAxe,
                                  chpSqlLib,codeJoin);

 codeChp:=code+chpSqlLib;
end;

//*****************************************************************************
//Objectif


function ChercheCodeSqlCaReelNewObj(const codeSession,structure,SaiCmdRef,SaiCmd,
         codeRestriction,chpClient,champclient:hstring;
         TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var j:integer;
    code,chpClientS,codeSql,codeE:hstring;
begin
 code:='';
 codeSql:='';
 codeE:='';
 for j:=1 to 10 do
  begin
   if TabNumValAxe[j]<>''
    then
     begin
      if code<>''
       then code:=code+' ,R.QBV_VALAXEP'+TabNumValAxe[j]
       else code:=' R.QBV_VALAXEP'+TabNumValAxe[j];
      codeE:= codeE+' AND C.QBR_VALAXENIV'+IntToStr(j)+
                    '=R.QBV_VALAXEP'+TabNumValAxe[j]+' ';
     end;
  end;


 chpClientS:=',R.QBV_VALAXEP'+chpClient;
 if SaiCmd=''
  then codeSql:='SELECT '+code+
                ',SUM(R.QBV_QTE1),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ',QBV_EXTRAPOLABLE FROM QBPPIVOT R '+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeRestriction+
                ' AND NOT EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
                'WHERE C.QBR_CODESESSION="'+codeSession+
                '" '+codeE+' AND '+champClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code+' ,QBV_EXTRAPOLABLE'
  else codeSql:='SELECT '+code+
                ',SUM(R.QBV_QTE1),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ',QBV_EXTRAPOLABLE FROM QBPPIVOT R '+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmd+
                '" '+codeRestriction+
                 ' AND NOT EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
                'WHERE C.QBR_CODESESSION="'+codeSession+
                '" '+codeE+' AND '+champClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code+' ,QBV_EXTRAPOLABLE';

 result:=codeSql;
end;

procedure RemplitCAReelNewObj(const codeSession,champclient,chpClient,SaiCmdRef,SaiCmd,structure,
          codeRestriction,chpSqlLib,JoinSqlLib:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR:array [1..11] of hstring;
    valP1,extra:double;
    valP2:integer;
begin
 //cherche le code du sql
 if not BPOkOrli
  then //codeSql:=ChercheCodeSqlCaClientVuPGI(nivMax,codeRestriction,chpClient,
       //                                     TabCodeAxe,DateDebPerC,DateFinPerC)
  else //-----------------> ORLI
  codeSql:=ChercheCodeSqlCaReelNewObj(codeSession,structure,SaiCmdRef,SaiCmd,
                                       codeRestriction,chpClient,
                                       champclient,TabNumValAxe,
                                       DateDebPerRef,DateFinPerRef,
                                      DateDebPerC,DateFinPerC);
  //ORLI <-----------------
 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAReelNewObj).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax].asFloat;
   valP2:=Q.fields[nivMax+1].asInteger;
   extra:=VALEUR(Q.fields[nivMax+2].asString);

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                   TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                   TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                   TabCodeAxeR[9],TabCodeAxeR[10]],
                   'CAREELNONVU',valP1);

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10]],
                      'NBCLIENTREELNONVU',
                      (valP2));
   QTBP.majValeur([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10]],
                      [0,0,0,0,0,0,0,0,0,0,0,
                      extra]);   
   Q.next;
  end;
 ferme(Q);
end;

function ChercheCodeSqlCaReelVuObj(const codeSession,structure,SaiCmdRef,SaiCmd,
         codeRestriction,chpClient,champclient:hstring;
         TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var j:integer;
    code,chpClientS,codeSql,codeE:hstring;
begin
 code:='';
 codeSql:='';
 codeE:='';
 for j:=1 to 10 do
  begin
   if TabNumValAxe[j]<>''
    then
     begin
      if code<>''
       then code:=code+' ,R.QBV_VALAXEP'+TabNumValAxe[j]
       else code:=' R.QBV_VALAXEP'+TabNumValAxe[j];
      codeE:= codeE+' AND C.QBR_VALAXENIV'+IntToStr(j)+
                    '=R.QBV_VALAXEP'+TabNumValAxe[j]+' ';
     end;
  end;


 chpClientS:=',R.QBV_VALAXEP'+chpClient;
 if SaiCmd=''
  then codeSql:='SELECT '+code+
                ',SUM(R.QBV_QTE1),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ',QBV_EXTRAPOLABLE FROM QBPPIVOT R '+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
                'WHERE C.QBR_CODESESSION="'+codeSession+
                '" '+codeE+' AND '+champClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code+' ,QBV_EXTRAPOLABLE'
  else codeSql:='SELECT '+code+
                ',SUM(R.QBV_QTE1),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ',QBV_EXTRAPOLABLE FROM QBPPIVOT R '+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmd+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
                'WHERE C.QBR_CODESESSION="'+codeSession+
                '" '+codeE+' AND '+champClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code+' ,QBV_EXTRAPOLABLE';

 result:=codeSql;
end;

procedure RemplitCAReelVuObj(const codeSession,champclient,chpClient,SaiCmdRef,SaiCmd,structure,
          codeRestriction,chpSqlLib,JoinSqlLib:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR:array [1..11] of hstring;
    valP1,extra:double;
    valP2:integer;
begin
 //cherche le code du sql
 if not BPOkOrli
  then// codeSql:=ChercheCodeSqlCaClientVuPGI(nivMax,codeRestriction,chpClient,
      //                                      TabCodeAxe,DateDebPerC,DateFinPerC)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaReelVuObj(codeSession,structure,SaiCmdRef,SaiCmd,
                                       codeRestriction,chpClient,
                                       champclient,TabNumValAxe,
                                       DateDebPerRef,DateFinPerRef,
                                       DateDebPerC,DateFinPerC);
       //ORLI <-----------------
 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAReelVuObj).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax].asFloat;
   valP2:=Q.fields[nivMax+1].asInteger;
   extra:=VALEUR(Q.fields[nivMax+2].asString);

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                   TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                   TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                   TabCodeAxeR[9],TabCodeAxeR[10]],
                   'CAREELVU',valP1);

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10]],
                      'NBCLIENTVU',
                      (valP2));
   QTBP.majValeur([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10]],
                      [0,0,0,0,0,0,0,0,0,0,0,
                      extra]);
   Q.next;
  end;
 ferme(Q);
end;

procedure MetArbreEnLigneObjectifHistoVu(NivMaxProfil:integer;
          const session,structure,chpClient,champClient,codeE:hstring;
          DateDebPerC,DateFinPerC:TDateTime;
          TabCodeAxe,TabNumValAxe:array of hstring;QTBP:TQTob);
var Q:TQuery;
    codesql:hstring;
    i:integer;
    TabValAxe:array [1..12] of hstring;
    newval:double;
    codewhere,groupby:hstring;
begin
 codesql:='SUM(A.QBR_REF1) ';
 groupby:='';
 for i:=1 to NivMaxProfil-1 do
  begin
   codesql:=codesql+',A.QBR_VALAXENIV'+IntToStr(i);
   if groupby=''
    then groupby:='A.QBR_VALAXENIV'+IntToStr(i)
    else groupby:=groupby+',A.QBR_VALAXENIV'+IntToStr(i);
  end;

 codewhere:=' AND EXISTS (SELECT C.QBV_QTE1 FROM QBPPIVOT C '+
            'WHERE C.QBV_CODESTRUCT="'+structure+
            '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
            '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
            '" '+codeE+' AND C.QBV_VALAXEP'+chpClient+
            '='+champClient+')';

 Q:=MOPenSql('SELECT '+codesql+' FROM QBPARBRE A '+
             ' WHERE A.QBR_CODESESSION="'+session+
             '" AND A.QBR_NIVEAU="'+IntToStr(NivMaxProfil)+
             '" group by '+groupby,//+codeWhere,
             'BPCoeff (MetArbreEnLigneObjectifHistoVu).',true);
 while not Q.eof do
  begin
   newval:=Q.fields[0].asFloat;
   for i:=1 to NivMaxProfil-1 do
    TabValAxe[i]:=Q.fields[i].asString;

  QTBP.majValeurChp([TabValAxe[1],TabValAxe[2],TabValAxe[3],TabValAxe[4],
                      TabValAxe[5],TabValAxe[6],TabValAxe[7],TabValAxe[8],
                      TabValAxe[9],TabValAxe[10]],
                      'CAHISTO',
                      (newval));

   Q.next;
  end;
 ferme(Q);

 Q:=MOPenSql('SELECT '+codesql+' FROM QBPARBRE A'+
             ' WHERE A.QBR_CODESESSION="'+session+
             '" AND A.QBR_NIVEAU="'+IntToStr(NivMaxProfil)+
             '" AND EXISTS (SELECT X.QBA_CODEAXE FROM QBPAXE X WHERE '+
             'X.QBA_CODEAXE="CLIENT" AND X.QBA_VALEURAXE='+champClient+
             ' AND X.QBA_CATEGORIEC="P") GROUP BY '+groupby,
             'BPCoeff (MetArbreEnLigneObjectifHistoVu).',true);
 while not Q.eof do
  begin
   newval:=Q.fields[0].asFloat;
   for i:=1 to NivMaxProfil-1 do
    TabValAxe[i]:=Q.fields[i].asString;

  QTBP.majValeurChp([TabValAxe[1],TabValAxe[2],TabValAxe[3],TabValAxe[4],
                      TabValAxe[5],TabValAxe[6],TabValAxe[7],TabValAxe[8],
                      TabValAxe[9],TabValAxe[10]],
                      'CAPROSPECT',
                      (newval));
   Q.next;
  end;
 ferme(Q);

 Q:=MOPenSql('SELECT '+codesql+' FROM QBPARBRE A'+
             ' WHERE A.QBR_CODESESSION="'+session+
             '" AND A.QBR_NIVEAU="'+IntToStr(NivMaxProfil)+
             '" '+codeWhere+' group by '+groupby,
             'BPCoeff (MetArbreEnLigneObjectifHistoVu).',true);
 while not Q.eof do
  begin
   newval:=Q.fields[0].asFloat;
   for i:=1 to NivMaxProfil-1 do
    TabValAxe[i]:=Q.fields[i].asString;

  QTBP.majValeurChp([TabValAxe[1],TabValAxe[2],TabValAxe[3],TabValAxe[4],
                      TabValAxe[5],TabValAxe[6],TabValAxe[7],TabValAxe[8],
                      TabValAxe[9],TabValAxe[10]],
                      'CAHISTOVU',
                      (newval));
   Q.next;
  end;
 ferme(Q);
end;

//*****************************************************************************
//Reel non vu

function ChercheCodeSqlCaClientNonVuPGI(niveauI:integer;
         codeRestriction,codeRestrictionL1,codeRestrictionL2,
         chpClient:hstring;
         TabCodeAxe:array of hstring;
         DateDebPerC,DateFinPerC,DateDebPerRef,DateFinPerRef:TDateTime):hstring;
var champs,joins:hstring;
begin
 DonneCodeChpJoinPGINbNiv(TabCodeAxe,niveauI,champs,joins);

 result:='SELECT '+champs+',SUM(L1.GL_TOTALHT),0 '+
         ' FROM TIERS T,LIGNE L1 '+joins+
         ' WHERE L1.GL_DATEPIECE>="'+USDATETIME(DateDebPerC)+
         '" AND L1.GL_DATEPIECE<="'+USDATETIME(DateFinPerC)+
         '" AND T.T_TIERS=L1.GL_TIERS AND T.T_NATUREAUXI="CLI" '+
         codeRestrictionL1+
         ' AND NOT EXISTS (SELECT L2.GL_TIERS FROM LIGNE L2 '+
         'WHERE  AND L1.GL_TIERS=L2.GL_TIERS '+
         ' AND L2.GL_DATEPIECE>="'+USDATETIME(DateDebPerRef)+
         '" AND L2.GL_DATEPIECE<="'+USDATETIME(DateFinPerRef)+
         '" '+codeRestrictionL2+') '+
         ' GROUP BY '+champs;
end;


function ChercheCodeSqlCaReelNonVu(const codeSession,structure,SaiCmdRef,SaiCmd,
         codeRestriction,chpClient:hstring;
         NivMax:integer;
         TabCodeAxe,TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var code,chpClientS,codeSql,codeE,ChpSum,codeJoin,codeRestrictionC,CodeJoinR:hstring;
begin
 codeSql:='';
 ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe,nivMax,
                            code,codeJoin,codeE);

 ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='1'
  then ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then ChpSum:='R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4+R.QBV_QTET5+'+
               'R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8+R.QBV_QTET9+R.QBV_QTET10+'+
               'R.QBV_QTET11+R.QBV_QTET12+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+'+
               'R.QBV_QTET16+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20';
 if DonneParamS(ps_BPCoeffPerCAQte)='3'
  then ChpSum:='R.QBV_QTE2';
 if DonneParamS(ps_BPCoeffPerCAQte)='4'
  then ChpSum:='R.QBV_QTE3';

 codeRestrictionC:=ChercheCodeRestrictionSession('C.','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);

 chpClientS:=',R.QBV_VALAXEP'+chpClient;
 if SaiCmd=''
  then codeSql:='SELECT '+code+
                ',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ' FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeRestriction+
                ' AND NOT EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code
  else codeSql:='SELECT '+code+
                ',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ' FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmd+
                '" '+codeRestriction+
                ' AND NOT EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code;


 result:=codeSql;
end;

procedure RemplitCAReelNonVu(const codeSession,chpClient,SaiCmdRef,SaiCmd,structure,
          codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR,TabLibAxeR:array [1..11] of hstring;
    valP1:double;
    valP2:integer;
begin
 //cherche le code du sql
 if not BPOkOrli
  then codeSql:=ChercheCodeSqlCaClientNonVuPGI(nivMax,codeRestriction,
                                            codeRestrictionL1,codeRestrictionL2,
                                            chpClient,
                                            TabCodeAxe,
                                            DateDebPerC,DateFinPerC,
                                            DateDebPerRef,DateFinPerRef)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaReelNonVu(codeSession,structure,SaiCmdRef,SaiCmd,
                                       codeRestriction,chpClient,
                                       NivMAx,TabCodeAxe,TabNumValAxe,
                                       DateDebPerRef,DateFinPerRef,
                                       DateDebPerC,DateFinPerC);
       //ORLI <-----------------
 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAReelNonVu).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to (nivMax*2)-1 do
    TabLibAxeR[j+1-nivMax]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax*2].asFloat;
   valP2:=Q.fields[(nivMax*2)+1].asInteger;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                   TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                   TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                   TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                   'CAREELNONVU',valP1);
   
   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'NBCLIENTREELNONVU',
                      (valP2));
   Q.next;
  end;
 ferme(Q);
end;

//*****************************************************************************
//Reel vu

function ChercheCodeSqlCaClientVuPGI(niveauI:integer;
         codeRestriction,codeRestrictionL1,codeRestrictionL2,
         chpClient:hstring;
         TabCodeAxe:array of hstring;
         DateDebPerC,DateFinPerC,DateDebPerRef,DateFinPerRef:TDateTime):hstring;
var champs,joins:hstring;
begin
 DonneCodeChpJoinPGINbNiv(TabCodeAxe,niveauI,champs,joins);

 result:='SELECT '+champs+',SUM(L1.GL_TOTALHT),0 '+
         ' FROM TIERS T,LIGNE L1 '+joins+
         ' WHERE L1.GL_DATEPIECE>="'+USDATETIME(DateDebPerC)+
         '" AND L1.GL_DATEPIECE<="'+USDATETIME(DateFinPerC)+
         '" AND T.T_TIERS=L1.GL_TIERS AND T.T_NATUREAUXI="CLI" '+
         codeRestrictionL1+
         ' AND EXISTS (SELECT L2.GL_TIERS FROM LIGNE L2 '+
         'WHERE  AND L1.GL_TIERS=L2.GL_TIERS '+
         ' AND L2.GL_DATEPIECE>="'+USDATETIME(DateDebPerRef)+
         '" AND L2.GL_DATEPIECE<="'+USDATETIME(DateFinPerRef)+
         '" '+codeRestrictionL2+') '+
         ' GROUP BY '+champs;
end;


function ChercheCodeSqlCaReelVu(const codeSession,structure,SaiCmdRef,SaiCmd,
         codeRestriction,chpClient:hstring;
         NivMax:integer;
         TabCodeAxe,TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var code,chpClientS,codeSql,codeE,ChpSum,codeJoin,codeRestrictionC,CodeJoinR:hstring;
begin
 code:='';
 ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe,nivMax,
                            code,codeJoin,codeE);

ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='1'
  then ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then ChpSum:='R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4+R.QBV_QTET5+'+
               'R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8+R.QBV_QTET9+R.QBV_QTET10+'+
               'R.QBV_QTET11+R.QBV_QTET12+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+'+
               'R.QBV_QTET16+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20';
 if DonneParamS(ps_BPCoeffPerCAQte)='3'
  then ChpSum:='R.QBV_QTE2';
 if DonneParamS(ps_BPCoeffPerCAQte)='4'
  then ChpSum:='R.QBV_QTE3';


 codeRestrictionC:=ChercheCodeRestrictionSession('C.','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);

 chpClientS:=',R.QBV_VALAXEP'+chpClient;
 if SaiCmd=''
  then codeSql:='SELECT '+code+
                ',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ' FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code
  else codeSql:='SELECT '+code+
                ',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                ' FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmd+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code;


 result:=codeSql;
end;

procedure RemplitCAReelVu(const codeSession,chpClient,SaiCmdRef,SaiCmd,structure,
          codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR,TabLibAxeR:array [1..11] of hstring;
    valP1:double;
    valP2:integer;
begin
 //cherche le code du sql
 if not BPOkOrli
  then codeSql:=ChercheCodeSqlCaClientVuPGI(nivMax,codeRestriction,
                                            codeRestrictionL1,codeRestrictionL2,
                                            chpClient,
                                            TabCodeAxe,DateDebPerC,DateFinPerC,
                                            DateDebPerRef,DateFinPerRef)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaReelVu(codeSession,structure,SaiCmdRef,SaiCmd,
                                       codeRestriction,chpClient,NivMax,
                                       TabCodeAxe,TabNumValAxe,
                                       DateDebPerRef,DateFinPerRef,
                                       DateDebPerC,DateFinPerC);
       //ORLI <-----------------
 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAReelVu).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to (nivMax*2)-1 do
    TabLibAxeR[j+1-nivMax]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax*2].asFloat;
   valP2:=Q.fields[(nivMax*2)+1].asInteger;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                   TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                   TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                   TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                   'CAREELVU',valP1);
   
   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'NBCLIENTVU',
                      (valP2));
   Q.next;
  end;
 ferme(Q);
end;

//*****************************************************************************
//Histo non vu

function ChercheCodeSqlCaHistoNonVuPgi(niveauI:integer;
         codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
         TabCodeAxe:array of hstring;
         DateCDeb,DateCFin,DateRefDeb,DateRefFin:TDateTime):hstring;
var champs,joins:hstring;
begin
 DonneCodeChpJoinPGINbNivAs(TabCodeAxe,niveauI,champs,joins);
 result:='SELECT '+champs+',SUM(L1.GL_TOTALHT),0 '+
         ' FROM TIERS T,LIGNE L1 '+joins+
         ' WHERE L1.GL_DATEPIECE>="'+USDATETIME(DateRefDeb)+
         '" AND L1.GL_DATEPIECE<="'+USDATETIME(DateRefFin)+
         '" AND T.T_TIERS=L1.GL_TIERS AND T.T_NATUREAUXI="CLI" '+
         codeRestrictionL1+
         ' AND EXISTS (SELECT L2.GL_TIERS FROM LIGNE L2 '+
         'WHERE L1.GL_TIERS=L2.GL_TIERS '+
         ' AND L2.GL_DATEPIECE>="'+USDATETIME(DateCDeb)+
         '" AND L2.GL_DATEPIECE<="'+USDATETIME(DateCFin)+
         '" AND L2.GL_TOTALHT=0 '+codeRestrictionL2+') '+
         ' GROUP BY '+champs;

end;


function ChercheCodeSqlCaHistoNonVu(const chpClient,structure,SaiCmdRef,SaiCmdC,
         codeRestriction:hstring;
         nivMax:integer;
         TabCodeAxe,TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var codeSql,code,codeE,ChpSum,codeJoin:hstring;
begin
 codeSql:='';
 ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe,nivMax,
                            code,codeJoin,codeE);

ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='1'
  then ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then ChpSum:='R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4+R.QBV_QTET5+'+
               'R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8+R.QBV_QTET9+R.QBV_QTET10+'+
               'R.QBV_QTET11+R.QBV_QTET12+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+'+
               'R.QBV_QTET16+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20';
 if DonneParamS(ps_BPCoeffPerCAQte)='3'
  then ChpSum:='R.QBV_QTE2';
 if DonneParamS(ps_BPCoeffPerCAQte)='4'
  then ChpSum:='R.QBV_QTE3';


 if SaiCmdRef=''
  then codeSql:='SELECT '+code+',SUM('+ChpSum+') ,COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT QBP_BPSAICMD FROM QBPCLTPERDU WHERE '+
                'QBP_BPSAICMD=R.QBV_VALAXEP8 AND QBP_BPDIVCOM=R.QBV_VALAXEP9 '+
                'AND QBP_BPCLIENT=R.QBV_VALAXEP13 AND QBP_BPSOCIETE=R.QBV_VALAXEP12) '+
                ' GROUP BY '+code
  else codeSql:='SELECT '+code+',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT QBP_BPSAICMD FROM QBPCLTPERDU WHERE '+
                'QBP_BPSAICMD=R.QBV_VALAXEP8 AND QBP_BPDIVCOM=R.QBV_VALAXEP9 '+
                'AND QBP_BPCLIENT=R.QBV_VALAXEP13 AND QBP_BPSOCIETE=R.QBV_VALAXEP12) '+
                ' GROUP BY '+code;
{  non vus
if SaiCmdRef=''
  then codeSql:='SELECT '+code+',SUM('+ChpSum+') ,COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeE+' AND C.QBV_VALAXEP'+chpClient+
                '=R.QBV_VALAXEP'+chpClient+
                ' AND (C.QBV_QTET1+C.QBV_QTET2+C.QBV_QTET3+C.QBV_QTET4+'+
                'C.QBV_QTET5+C.QBV_QTET6+C.QBV_QTET7+C.QBV_QTET8+'+
                'C.QBV_QTET9+C.QBV_QTET10+C.QBV_QTET11+C.QBV_QTET12+'+
                'C.QBV_QTET13+C.QBV_QTET14+C.QBV_QTET15+C.QBV_QTET16+'+
                'C.QBV_QTET17+C.QBV_QTET18+C.QBV_QTET19+C.QBV_QTET20)=0) '+
                ' GROUP BY '+code
  else codeSql:='SELECT '+code+',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_VALAXEP8="'+SaiCmdC+
                '" '+codeE+' AND C.QBV_VALAXEP'+chpClient+
                '=R.QBV_VALAXEP'+chpClient+
                ' AND (C.QBV_QTET1+C.QBV_QTET2+C.QBV_QTET3+C.QBV_QTET4+'+
                'C.QBV_QTET5+C.QBV_QTET6+C.QBV_QTET7+C.QBV_QTET8+'+
                'C.QBV_QTET9+C.QBV_QTET10+C.QBV_QTET11+C.QBV_QTET12+'+
                'C.QBV_QTET13+C.QBV_QTET14+C.QBV_QTET15+C.QBV_QTET16+'+
                'C.QBV_QTET17+C.QBV_QTET18+C.QBV_QTET19+C.QBV_QTET20)=0) '+
                ' GROUP BY '+code;
/}
//attention si temps calcul un peu long faire un test avec NOT IN
//à la place de NOT EXISTS
 result:=codeSql;
end;

procedure RemplitCAHistoNonVu(const chpClient,structure,SaiCmdRef,SaiCmdC,
          codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef:TDateTime;
          DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR,TabLibAxeR:array [1..11] of hstring;
    valP1:double;
    valP2:integer;
begin
 //cherche le code du sql
 //càd les champs qui faut prendre dans la table pivot
 if not BPOkOrli
  then codeSql:=ChercheCodeSqlCaHistoNonVuPgi(nivMax,codeRestriction,
                                         codeRestrictionL1,codeRestrictionL2,
                                         TabCodeAxe,
                                         DateDebPerC,DateFinPerC,
                                         DateDebPerRef,DateFinPerRef)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaHistoNonVu(chpClient,structure,
                                      SaiCmdRef,SaiCmdC,codeRestriction,
                                      nivMax,TabCodeAxe,TabNumValAxe,
                                      DateDebPerRef,DateFinPerRef,
                                      DateDebPerC,DateFinPerC);
       //ORLI <-----------------

 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAHistoNonVu).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to (nivMax*2)-1 do
    TabLibAxeR[j+1-nivMax]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax*2].asFloat;
   valP2:=Q.fields[(nivMax*2)+1].asInteger;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'CAHISTONONVU',
                      (valP1));
    QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'NBCLIENTHISTONONVU',
                      (valP2));
   Q.next;
  end;
 ferme(Q);
end;

//*****************************************************************************
//Histo vu

function ChercheCodeSqlCaHistoVuPgi(niveauI:integer;
         codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
         TabCodeAxe:array of hstring;
         DateCDeb,DateCFin,DateRefDeb,DateRefFin:TDateTime):hstring;
var champs,joins:hstring;
begin
 DonneCodeChpJoinPGINbNivAs(TabCodeAxe,niveauI,champs,joins);
 result:='SELECT '+champs+',SUM(L1.GL_TOTALHT),0 '+
         ' FROM TIERS T,LIGNE L1 '+joins+
         ' WHERE L1.GL_DATEPIECE>="'+USDATETIME(DateRefDeb)+
         '" AND L1.GL_DATEPIECE<="'+USDATETIME(DateRefFin)+
         '" AND T.T_TIERS=L1.GL_TIERS AND T.T_NATUREAUXI="CLI" '+
         codeRestrictionL1+
         ' AND EXISTS (SELECT L2.GL_TIERS FROM LIGNE L2 '+
         'WHERE  AND L1.GL_TIERS=L2.GL_TIERS '+
         ' AND L2.GL_DATEPIECE>="'+USDATETIME(DateCDeb)+
         '" AND L2.GL_DATEPIECE<="'+USDATETIME(DateCFin)+
         '" '+codeRestrictionL2+') '+
         ' GROUP BY '+champs;

end;


function ChercheCodeSqlCaHisto(const codeSession,chpClient,structure,SaiCmdRef,SaiCmdC,
         codeRestriction:hstring;
         nivMax:integer;
         TabCodeAxe,TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var codeSql,code,codeE,ChpSum,codeJoin,codeRestrictionC,CodeJoinR:hstring;
begin
 codeSql:='';
 ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe,nivMax,
                            code,codeJoin,codeE);
ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='1'
  then ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then ChpSum:='R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4+R.QBV_QTET5+'+
               'R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8+R.QBV_QTET9+R.QBV_QTET10+'+
               'R.QBV_QTET11+R.QBV_QTET12+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+'+
               'R.QBV_QTET16+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20';
 if DonneParamS(ps_BPCoeffPerCAQte)='3'
  then ChpSum:='R.QBV_QTE2';
 if DonneParamS(ps_BPCoeffPerCAQte)='4'
  then ChpSum:='R.QBV_QTE3';


 codeRestrictionC:=ChercheCodeRestrictionSession('C.','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);

 if SaiCmdRef=''
  then codeSql:='SELECT '+code+',SUM('+ChpSum+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code
  else codeSql:='SELECT '+code+',SUM('+ChpSum+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeRestriction+
                ' AND EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_VALAXEP8="'+SaiCmdC+
                '" '+codeE+codeRestrictionC+
                ' AND C.QBV_VALAXEP'+chpClient+'=R.QBV_VALAXEP'+chpClient+')'+
                ' GROUP BY '+code;

 result:=codeSql;
end;

procedure RemplitCAHistoVu(const codeSession,chpClient,structure,SaiCmdRef,SaiCmdC,
          codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef:TDateTime;
          DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR,TabLibAxeR:array [1..11] of hstring;
    valP1:double;
begin
 //cherche le code du sql
 //càd les champs qui faut prendre dans la table pivot
 if not BPOkOrli
  then codeSql:=ChercheCodeSqlCaHistoVuPgi(nivMax,codeRestriction,
                                         codeRestrictionL1,codeRestrictionL2,
                                         TabCodeAxe,
                                         DateDebPerC,DateFinPerC,
                                         DateDebPerRef,DateFinPerRef)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaHisto(codeSession,chpClient,structure,
                                      SaiCmdRef,SaiCmdC,codeRestriction,
                                      nivMax,TabCodeAxe,TabNumValAxe,
                                      DateDebPerRef,DateFinPerRef,
                                      DateDebPerC,DateFinPerC);
       //ORLI <-----------------

 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAHistoVu).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to (nivMax*2)-1 do
    TabLibAxeR[j+1-nivMax]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax*2].asFloat;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'CAHISTOVU',
                      (valP1));
   Q.next;
  end;
 ferme(Q);
end;

//****************************************************************************
//histo

function ChercheCodeSqlCaHistoT(OkObj:boolean;
         const chpClient,structure,SaiCmdRef,SaiCmdC,
         codeRestriction:hstring;
         nivMax:integer;
         TabCodeAxe,TabNumValAxe:array of hstring;
         DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime):hstring;
var codeSql,code,codeE,ChpSum,codeJoin:hstring;
begin
 codeSql:='';
 ChercheCodeSqlChpsEtChpLib(TabCodeAxe,TabNumValAxe,nivMax,
                            code,codeJoin,codeE);


 ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='1'
  then ChpSum:='R.QBV_QTE1';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then ChpSum:='R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4+R.QBV_QTET5+'+
               'R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8+R.QBV_QTET9+R.QBV_QTET10+'+
               'R.QBV_QTET11+R.QBV_QTET12+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+'+
               'R.QBV_QTET16+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20';
 if DonneParamS(ps_BPCoeffPerCAQte)='3'
  then ChpSum:='R.QBV_QTE2';
 if DonneParamS(ps_BPCoeffPerCAQte)='4'
  then ChpSum:='R.QBV_QTE3';

 if SaiCmdRef=''
  then codeSql:='SELECT '+code+',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerRef)+
                '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerRef)+
                '" '+codeRestriction+
                ' AND NOT EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebPerC)+
                '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinPerC)+
                '" '+codeE+' AND C.QBV_VALAXEP'+chpClient+
                '=R.QBV_VALAXEP'+chpClient+
                ') GROUP BY '+code
  else codeSql:='SELECT '+code+',SUM('+ChpSum+'),COUNT (DISTINCT R.QBV_VALAXEP'+chpClient+') '+
                'FROM QBPPIVOT R '+codeJoin+
                ' WHERE R.QBV_CODESTRUCT="'+structure+
                '" AND R.QBV_VALAXEP8="'+SaiCmdRef+
                '" '+codeRestriction+
                ' AND NOT EXISTS (SELECT C.QBV_QTE2 FROM QBPPIVOT C '+
                'WHERE C.QBV_CODESTRUCT="'+structure+
                '" AND C.QBV_VALAXEP8="'+SaiCmdC+
                '" '+codeE+' AND C.QBV_VALAXEP'+chpClient+
                '=R.QBV_VALAXEP'+chpClient+
                ') GROUP BY '+code;

 result:=codeSql;
end;

function ChercheCodeSqlCaHistoPgi(niveauI:integer;
         codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
         TabCodeAxe:array of hstring;
         DateRefDeb,DateRefFin:TDateTime):hstring;
var champs,joins:hstring;
begin
 DonneCodeChpJoinPGINbNivAs(TabCodeAxe,niveauI,champs,joins);
 result:='SELECT '+champs+',SUM(L1.GL_TOTALHT),0 '+
         ' FROM TIERS T,LIGNE L1 '+joins+
         ' WHERE L1.GL_DATEPIECE>="'+USDATETIME(DateRefDeb)+
         '" AND L1.GL_DATEPIECE<="'+USDATETIME(DateRefFin)+
         '" AND T.T_TIERS=L1.GL_TIERS AND T.T_NATUREAUXI="CLI" '+
         codeRestrictionL1+
         ' GROUP BY '+champs;
end;

procedure RemplitCAHisto(OkObj:boolean;
          const chpClient,structure,SaiCmdRef,SaiCmdC,
          codeRestriction,codeRestrictionL1,codeRestrictionL2:hstring;
          TabCodeAxe,TabNumValAxe:array of hstring;
          DateDebPerRef,DateFinPerRef:TDateTime;
          DateDebPerC,DateFinPerC:TDateTime;
          nivMax:integer;
          var QTBP:TQTob);
var Q:TQuery;
    codeSql:hstring;
    j:integer;
    TabCodeAxeR,TabLibAxeR:array [1..11] of hstring;
    valP1:double;
    valP2:integer;
begin
 //cherche le code du sql
 //càd les champs qui faut prendre dans la table pivot
 if not BPOkOrli
  then codeSql:=ChercheCodeSqlCaHistoPgi(nivMax,codeRestriction,
                                         codeRestrictionL1,codeRestrictionL2,
                                         TabCodeAxe,
                                         DateDebPerRef,DateFinPerRef)
  else //-----------------> ORLI
       codeSql:=ChercheCodeSqlCaHistoT(OkObj,chpClient,structure,
                                      SaiCmdRef,SaiCmdC,codeRestriction,
                                      nivMax,TabCodeAxe,TabNumValAxe,
                                      DateDebPerRef,DateFinPerRef,
                                      DateDebPerC,DateFinPerC);
       //ORLI <-----------------
                                             
 Q:=MOpenSql(codeSql,'BPCoeff (RemplitCAHisto).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to (nivMax*2)-1 do
    TabLibAxeR[j+1-nivMax]:=Q.fields[j].asString;
   valP1:=Q.fields[nivMax*2].asFloat;
   valP2:=Q.fields[(nivMax*2)+1].asInteger;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'CAHISTO',
                      (valP1));
   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibAxeR[1],TabLibAxeR[2],
                      TabLibAxeR[3],TabLibAxeR[4],TabLibAxeR[5],
                      TabLibAxeR[6],TabLibAxeR[7],TabLibAxeR[8],
                      TabLibAxeR[9],TabLibAxeR[10]],
                      'NBCLIENTHISTO',
                      (valP2));
   Q.next;
  end;
 ferme(Q);
end;

procedure InitialiseCoeffPrev(const codeSession:hstring);
var QTBP: TQTob;
    structure,SaiCmdRef,chpClient,SaiCmd,codeRestriction:hstring;
    codeRestrictionL1,codeRestrictionL2:hstring;
    TabNumValAxe,TabCodeAxe:array [1..11] of hstring;
    DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC:TDateTime;
    nivMax,j:integer;
    chpSqlLib,joinSqlLib,SessionObjectif,codeE,champClient,CodeJoin:hstring;
begin
 MExecuteSql('DELETE FROM QBPCOEFF '+
             'WHERE QBT_CODESESSION="'+codeSession+'"',
             'BPCoeff (InitialiseCoeffPrev).');
 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 if not BPOkOrli
  then
   begin
    chpClient:='GL_TIERS';
    structure:='';
    SaiCmd:='';
    SaiCmdRef:='';
    ChercheCodeAxeSession(codeSession,TabCodeAxe);
   end
  else
   //-----------------> ORLI
   begin
    chpClient:='13';
    structure:=ChercheSessionStructure(codeSession);
    ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);
    ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
   end;
   //ORLI <-----------------

 ChercheDateDDateFPeriode(codeSession,DateDebPerC,DateFinPerC,
                          DateDebPerRef,DateFinPerRef);

 //création de la tob
 if QTBP <> nil
  then
   begin
    QTBP.Free;
    QTBP:=nil;
   end;
 QTBP:=TQTob.create(['AXE1','AXE2','AXE3','AXE4','AXE5',
                     'AXE6','AXE7','AXE8','AXE9','AXE10',
                     'LIB1','LIB2','LIB3','LIB4','LIB5',
                     'LIB6','LIB7','LIB8','LIB9','LIB10'],
                    ['CAHISTOVU','CAHISTONONVU',
                     'CAREELVU','CAREELNONVU','COEFF','CAHISTO',
                     'NBCLIENTHISTO','NBCLIENTVU','NBCLIENTHISTONONVU',
                     'NBCLIENTREELNONVU','CAPROSPECT','EXTRA']);

 codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,CodeJoin);
 codeRestrictionL1:=ChercheCodeRestrictionSession('L1.','QBV_VALAXEP','QBV_',codeSession,CodeJoin);
 codeRestrictionL2:=ChercheCodeRestrictionSession('L2.','QBV_VALAXEP','QBV_',codeSession,CodeJoin);

 if not ChercheSessionObjectif(codeSession,SessionObjectif)
  then
   begin
    //libellé
    RemplitCAHisto(false,chpClient,structure,SaiCmdRef,SaiCmd,codeRestriction,
                     codeRestrictionL1,codeRestrictionL2,
                     TabCodeAxe,TabNumValAxe,
                     DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                     nivMax,QTBP);
    RemplitCAHistoVu(codeSession,chpClient,structure,SaiCmdRef,SaiCmd,codeRestriction,
                     codeRestrictionL1,codeRestrictionL2,
                     TabCodeAxe,TabNumValAxe,
                     DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                     nivMax,QTBP);
    RemplitCAHistoNonVu(chpClient,structure,SaiCmdRef,SaiCmd,codeRestriction,
                     codeRestrictionL1,codeRestrictionL2,
                     TabCodeAxe,TabNumValAxe,
                     DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                     nivMax,QTBP);
    RemplitCAReelVu(codeSession,chpClient,SaiCmdRef,SaiCmd,structure,codeRestriction,
                    codeRestrictionL1,codeRestrictionL2,
                    TabCodeAxe,TabNumValAxe,
                    DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                    nivMax,QTBP);
    RemplitCAReelNonVu(codeSession,chpClient,SaiCmdRef,SaiCmd,structure,codeRestriction,
                    codeRestrictionL1,codeRestrictionL2,
                    TabCodeAxe,TabNumValAxe,
                    DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                    nivMax,QTBP);
      end
     else
      begin
       codeE:='';
       for j:=1 to nivMax+1 do
        begin
         if TabNumValAxe[j]<>''
          then
           begin
            codeE:= codeE+' AND A.QBR_VALAXENIV'+IntToStr(j-1)+
                    '=C.QBV_VALAXEP'+TabNumValAxe[j]+' ';
           end;
        end;
       //CAS on part de OBJECTIF
       //cherche champ client dans arbre
       champClient:=ChercheClientArbreSession(SessionObjectif,nivMax+1);

       //cherche le CA histo (càd objectif)
       //le CA prospect et le CA histo vu
       MetArbreEnLigneObjectifHistoVu(NivMax+1,SessionObjectif,structure,
                                      chpClient,'A.'+champClient,codeE,
                                      DateDebPerC,DateFinPerC,
                                      TabCodeAxe,TabNumValAxe,QTBP);

       //recupère le réalisé dans qbppivot
       //le CA des clients vus
       RemplitCAReelVuObj(SessionObjectif,'C.'+champClient,chpClient,SaiCmdRef,SaiCmd,structure,codeRestriction,
                       chpSqlLib,joinSqlLib,
                       TabCodeAxe,TabNumValAxe,
                       DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                       nivMax,QTBP);
       //recupère le CA des clients nouveaux
       RemplitCAReelNewObj(SessionObjectif,'C.'+champClient,chpClient,SaiCmdRef,SaiCmd,structure,codeRestriction,
                       chpSqlLib,joinSqlLib,
                       TabCodeAxe,TabNumValAxe,
                       DateDebPerRef,DateFinPerRef,DateDebPerC,DateFinPerC,
                       nivMax,QTBP);
      end;

    RemplitTableQBP(codeSession,QTBP);

    MAJExtrapolable(codesession);

    if DonneParamB(pb_BPCoeffGene)
     then MiseAJourDesCoeffsAvecCoeffGene(codeSession);

    if QTBP <> nil
     then
      begin
       QTBP.Free;
       QTBP:=nil;
      end;

 //mise à jour dans qbpsessionbp session init
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="X",QBS_INITCOEFF="X" '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPCoeff (InitialiseCoeffPrev).');

end;

//*****************************************************************************
//calcul coeff permanent

procedure RemplitTableQBPPermanent(const codeSession:hstring;
          QTBP:TQTob);
var i:integer;
    TobF:Tob;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:hstring;
    lib1,lib2,lib3,lib4,lib5,lib6,lib7,lib8,lib9,lib10:hstring;
    cahisto,careel,coeff,cahistotot,careeltot:double;
begin
 for i := 0 to QTBP.laTob.Detail.Count - 1 do
  begin
   TobF := QTBP.laTob.Detail[i];
   axe1 := TobF.GetValue('AXE1');
   axe2 := TobF.GetValue('AXE2');
   axe3 := TobF.GetValue('AXE3');
   axe4 := TobF.GetValue('AXE4');
   axe5 := TobF.GetValue('AXE5');
   axe6 := TobF.GetValue('AXE6');
   axe7 := TobF.GetValue('AXE7');
   axe8 := TobF.GetValue('AXE8');
   axe9 := TobF.GetValue('AXE9');
   axe10 := TobF.GetValue('AXE10');
   lib1 := TobF.GetValue('LIBAXE1');
   lib2 := TobF.GetValue('LIBAXE2');
   lib3 := TobF.GetValue('LIBAXE3');
   lib4 := TobF.GetValue('LIBAXE4');
   lib5 := TobF.GetValue('LIBAXE5');
   lib6 := TobF.GetValue('LIBAXE6');
   lib7 := TobF.GetValue('LIBAXE7');
   lib8 := TobF.GetValue('LIBAXE8');
   lib9 := TobF.GetValue('LIBAXE9');
   lib10 := TobF.GetValue('LIBAXE10');
   cahisto := TobF.GetValue('CAHISTO');
   careel := TobF.GetValue('CAREEL');
   cahistotot := TobF.GetValue('CAHISTOTOT');
   careeltot := TobF.GetValue('CAREELTOT');
   if cahisto<>0
    then coeff:=careel/cahisto
    else coeff:=0;
   if careel=0
    then coeff:=0;
   MExecuteSql('INSERT INTO QBPCOEFF (QBT_CODESESSION,QBT_NUMCOEFF,'+
               'QBT_CODEAXEC1,QBT_CODEAXEC2,QBT_CODEAXEC3,QBT_CODEAXEC4,'+
               'QBT_CODEAXEC5,QBT_CODEAXEC6,QBT_CODEAXEC7,QBT_CODEAXEC8,'+
               'QBT_CODEAXEC9,QBT_CODEAXEC10,'+
               'QBT_LIBC1,QBT_LIBC2,QBT_LIBC3,QBT_LIBC4,'+
               'QBT_LIBC5,QBT_LIBC6,QBT_LIBC7,QBT_LIBC8,'+
               'QBT_LIBC9,QBT_LIBC10,QBT_CAHISTOVU,QBT_CAHISTONONVU,'+
               'QBT_CAREELVU,QBT_CAREELNEW,'+
               'QBT_COEFFEXTRAPO,QBT_CARETENU,QBT_COEFFRETENU,QBT_CAHISTO,'+
               'QBT_NBCLTHISTO,QBT_NBCLTVU,QBT_NBCLTHISTONVU,'+
               'QBT_NBCLTREELNONVU,QBT_CAPROSPECT,QBT_XVISIT,'+
               'QBT_NEWPROSPECT,QBT_RESTEAVISITER,QBT_EXTRAPOLABLE) ' +
               'VALUES ("'+codeSession+'","'+strFPoint4(i)+'","'+axe1+'","'+axe2+
               '","'+axe3+'","'+axe4+'","'+axe5+'","'+axe6+'","'+
               axe7+'","'+axe8+'","'+axe9+'","'+axe10+'","'+
               lib1+'","'+lib2+
               '","'+lib3+'","'+lib4+'","'+lib5+'","'+lib6+'","'+
               lib7+'","'+lib8+'","'+lib9+'","'+lib10+'","'+
               strFPoint4(cahistotot)+
               '","0","'+strFPoint4(careeltot)+
               '","0","'+strFPoint4(coeff)+'","0","'+strFPoint4(coeff)+
               '","'+strFPoint4(cahistotot)+
               '","0","0","0","0","0","0","0","0","0")',
               'BPCoeff (RemplitTableQBPPermanent).');
  end;
end;

//-----------------> ORLI
procedure ChercheCodeSqlCoeffOrli(structure,codeRestriction,chpSqlLib,joinSqlLib:hstring;
          DateDebCalcul,DateFinCalcul:TDateTime;
          TabNumValAxe:array of hstring;
          var codeSql:hstring);
var j:integer;
    code:hstring;
begin
 code:='';
 codeSql:='';
 for j:=1 to 10 do
  begin
   if TabNumValAxe[j]<>''
    then
     begin
      if code<>''
       then code:=code+' ,R.QBV_VALAXEP'+TabNumValAxe[j]
       else code:=' R.QBV_VALAXEP'+TabNumValAxe[j];
     end;
  end;

 codeSql:='SELECT '+code+chpSqlLib+',SUM(R.QBV_QTE2),'+
          'SUM(R.QBV_QTET1+R.QBV_QTET2+R.QBV_QTET3+R.QBV_QTET4'+
          '+R.QBV_QTET5+R.QBV_QTET6+R.QBV_QTET7+R.QBV_QTET8'+
          '+R.QBV_QTET9+R.QBV_QTET10+R.QBV_QTET11+R.QBV_QTET12'+
          '+R.QBV_QTET13+R.QBV_QTET14+R.QBV_QTET15+R.QBV_QTET16'+
          '+R.QBV_QTET17+R.QBV_QTET18+R.QBV_QTET19+R.QBV_QTET20) '+
          'FROM QBPPIVOT R '+joinSqlLib+
          ' WHERE R.QBV_CODESTRUCT="'+structure+
          '" AND R.QBV_DATEPIVOT>="'+USDATETIME(DateDebCalcul)+
          '" AND R.QBV_DATEPIVOT<="'+USDATETIME(DateFinCalcul)+
          '" '+codeRestriction+
          ' GROUP BY '+code+chpSqlLib;

end;
//ORLI <-----------------

procedure ChercheCodeSqlCoeffPgi(codeRestriction,chpSqlLib,joinSqlLib:hstring;
          DateDebCalcul,DateFinCalcul:TDateTime;
          var codeSql:hstring);
begin
 CodeSql:='';
end;


procedure ChercheCodeSqlCoeff(structure,codeRestriction,chpSqlLib,joinSqlLib:hstring;
          DateDebCalcul,DateFinCalcul:TDateTime;
          TabNumValAxe:array of hstring;
          var codeSql:hstring);
begin
 if BPOkOrli
  then //-----------------> ORLI
       ChercheCodeSqlCoeffOrli(structure,codeRestriction,chpSqlLib,joinSqlLib,
                               DateDebCalcul,DateFinCalcul,TabNumValAxe,
                               codeSql)
       //ORLI <-----------------
  else ChercheCodeSqlCoeffPgi(codeRestriction,chpSqlLib,joinSqlLib,
                              DateDebCalcul,DateFinCalcul,codeSql);
end;

function ChercheCodeSqlCoeffPermanent(
         structure,codeRestriction,chpSqlLib,joinSqlLib:hstring;
         TabNumValAxe:array of hstring;
         nivMax:integer;
         DateDebCalcul,DateFinCalcul:TDateTime;
         DateDebCourante,DateFinCourante:TDateTime;
         DateDebReference,DateFinReference:TDateTime;
         var QTBP:TQTob):hstring;
var codeSql,CoeffPerCAQte:hstring;
    j:integer;
    Q:TQuery;
    TabCodeAxeR,TabLibR:array [1..11] of hstring;
    CA:double;
    //qte:double;
begin
 CoeffPerCAQte:=DonneParamS(ps_BPCoeffPerCAQte);

 //histo
 ChercheCodeSqlCoeff(structure,codeRestriction,chpSqlLib,joinSqlLib,
                     DateDebCalcul-365,DateFinCalcul-365,TabNumValAxe,
                     codeSql);

 Q:=MOpenSql(codeSql,'BPCoeff (ChercheCodeSqlCoeffPermanent).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to nivMax+nivMax-1 do
    TabLibR[j+1-nivMax]:=Q.fields[j].asString;
   CA:=Q.fields[nivMax+nivMax].asFloat;
   //calacul du coeff à partir des qtes
   if CoeffPerCAQte='2'
    then CA:=Q.fields[nivMax+nivMax+1].asFloat;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibR[1],TabLibR[2],
                      TabLibR[3],TabLibR[4],TabLibR[5],
                      TabLibR[6],TabLibR[7],TabLibR[8],
                      TabLibR[9],TabLibR[10]],
                      'CAHISTO',
                      (CA));
   Q.next;
  end;
 ferme(Q);

 //ventes
 ChercheCodeSqlCoeff(structure,codeRestriction,chpSqlLib,joinSqlLib,
                     DateDebCalcul,DateFinCalcul,TabNumValAxe,
                     codeSql);

 Q:=MOpenSql(codeSql,'BPCoeff (ChercheCodeSqlCoeffPermanent).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to nivMax+nivMax-1 do
    TabLibR[j+1-nivMax]:=Q.fields[j].asString;
   CA:=Q.fields[nivMax+nivMax].asFloat;
   //calacul du coeff à partir des qtes
   if CoeffPerCAQte='2'
    then CA:=Q.fields[nivMax+nivMax+1].asFloat;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibR[1],TabLibR[2],
                      TabLibR[3],TabLibR[4],TabLibR[5],
                      TabLibR[6],TabLibR[7],TabLibR[8],
                      TabLibR[9],TabLibR[10]],
                      'CAREEL',
                      (CA));
   Q.next;
  end;
 ferme(Q);


 //histo total
 ChercheCodeSqlCoeff(structure,codeRestriction,chpSqlLib,joinSqlLib,
                     DateDebReference,DateFinReference,TabNumValAxe,
                     codeSql);

 Q:=MOpenSql(codeSql,'BPCoeff (ChercheCodeSqlCoeffPermanent).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to nivMax+nivMax-1 do
    TabLibR[j+1-nivMax]:=Q.fields[j].asString;
   CA:=Q.fields[nivMax+nivMax].asFloat;
   //calcul du coeff à partir des qtes
   if CoeffPerCAQte='2'
    then CA:=Q.fields[nivMax+nivMax+1].asFloat;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibR[1],TabLibR[2],
                      TabLibR[3],TabLibR[4],TabLibR[5],
                      TabLibR[6],TabLibR[7],TabLibR[8],
                      TabLibR[9],TabLibR[10]],
                      'CAHISTOTOT',
                      (CA));
   Q.next;
  end;
 ferme(Q);

 //ventes totales
  ChercheCodeSqlCoeff(structure,codeRestriction,chpSqlLib,joinSqlLib,
                     DateDebCourante,DateFinCourante,TabNumValAxe,
                     codeSql);

 Q:=MOpenSql(codeSql,'BPCoeff (ChercheCodeSqlCoeffPermanent).',true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    TabCodeAxeR[j+1]:=Q.fields[j].asString;
   for j:=nivMax to nivMax+nivMax-1 do
    TabLibR[j+1-nivMax]:=Q.fields[j].asString;
   CA:=Q.fields[nivMax+nivMax].asFloat;
   //calacul du coeff à partir des qtes
   if CoeffPerCAQte='2'
    then CA:=Q.fields[nivMax+nivMax+1].asFloat;

   QTBP.majValeurChp([TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibR[1],TabLibR[2],
                      TabLibR[3],TabLibR[4],TabLibR[5],
                      TabLibR[6],TabLibR[7],TabLibR[8],
                      TabLibR[9],TabLibR[10]],
                      'CAREELTOT',
                      (CA));
   Q.next;
  end;
 ferme(Q);
end;

//dans le cas des articles permanents
//on calculera une tendance par rapport à une période  paramétrée
// de 2 3 ou 6 mois
procedure InitialiseCoeffTendancePrev(const codeSession:hstring);
var DateDebCalcul,DateFinCalcul:TDateTime;
    nivMax:integer;
    structure,codeRestriction,chpSqlLib,joinSqlLib:hstring;
    TabNumValAxe,TabCodeAxe:array [1..11] of hstring;
    QTBP:TQTob;
    DateDebCourante,DateFinCourante:TDateTime;
    DateDebReference,DateFinReference:TDateTime;
    champs,joinChpSql,CodeJoinR:hstring;
begin
 //recupère période de calcul
 DateDebCalcul:=date-DonneParamI(pi_BPNbJourAvantJ);
 DateFinCalcul:=date+DonneParamI(pi_BPNbJourApresJ);

 //calcul coeff
 //pour les axes de la session

 MExecuteSql('DELETE FROM QBPCOEFF '+
            'WHERE QBT_CODESESSION="'+codeSession+'"',
            'BPCoeff (InitialiseCoeffTendancePrev).');
 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(codeSession);

 structure:=ChercheSessionStructure(codeSession);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);


 codeRestriction:=ChercheCodeRestrictionSession('R.','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);
 ChercheDateDDateFPeriode(codeSession,DateDebCourante,DateFinCourante,
                          DateDebReference,DateFinReference);

 //création de la tob
 QTBP:=TQTob.create(['AXE1','AXE2','AXE3','AXE4','AXE5',
                     'AXE6','AXE7','AXE8','AXE9','AXE10',
                     'LIBAXE1','LIBAXE2','LIBAXE3','LIBAXE4','LIBAXE5',
                     'LIBAXE6','LIBAXE7','LIBAXE8','LIBAXE9','LIBAXE10'],
                    ['CAHISTO','CAREEL','COEFF','CAHISTOTOT','CAREELTOT']);

 champs:='';
 joinChpSql:='';

 if BPOkOrli
  then //-----------------> ORLI
       ChercheCodeSqlChpAxeLibAvecPivot(NivMax,'R.',TabCodeAxe,TabNumValAxe,
                                  chpSqlLib,joinSqlLib)
       //ORLI <-----------------
  else
   begin        
    DonneCodeChpJoinPGINbNiv(TabCodeAxe,nivMax,champs,joinChpSql);
    ChercheCodeSqlChpAxeLibPGI(nivMax,false,TabCodeAxe,'',chpSqlLib,joinSqlLib);
   end;

 ChercheCodeSqlCoeffPermanent(structure,codeRestriction,
                              champs+chpSqlLib,joinChpSql+joinSqlLib,TabNumValAxe,
                              nivMax,DateDebCalcul,DateFinCalcul,
                              DateDebCourante,DateFinCourante,
                              DateDebReference,DateFinReference,QTBP);

 RemplitTableQBPPermanent(codeSession,QTBP);

 if QTBP <> nil
  then
   begin
    QTBP.Free;
    QTBP:=nil;
   end;

 //mise à jour dans qbpsessionbp session init
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="X" '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPCoeff (InitialiseCoeffTendancePrev).');
end;

//*****************************************************************************
//Calcul coeff
{
procedure CalculCoeff(const codeSession:string);
begin
 if OkSessionPermanent(codeSession)
  then InitialiseCoeffTendancePrev(codeSession)
  else InitialiseCoeffPrev(codeSession);
end;  }

procedure CalculCoeff(const codeSession:hstring);
var i:integer;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:hstring;
    lib1,lib2,lib3,lib4,lib5,lib6,lib7,lib8,lib9,lib10,extra:hstring;
    cahistovu,cahistononvu,careelvu,careelnonvu,coeff,cahisto,CAProspect:double;
    Xvisit,NewProspect,ResteAVisiter,caRetenu:double;
    nbCltHisto,nbCltVu,nbCltHistoNonVu,nbCltHistoReelNonVu:integer;
    PrctSeuilCaHVu,PrctCaHVu,CAReelVux:double;
    Q:TQUery;
begin
 MExecuteSql('DELETE FROM QBPCOEFF '+
             'WHERE QBT_CODESESSION="'+codeSession+'"',
             'BPCoeff (InitialiseCoeffPrev).');

 RemplitTableQBPCalculTmp(codeSession);
 i:=1;
 PrctSeuilCaHVu:=DonneParamR(pr_BPPrctSeuilCaHVu);
 Q:=MOPenSql('SELECT QBU_VALAXECU1,QBU_VALAXECU2,QBU_VALAXECU3,QBU_VALAXECU4,'+
             'QBU_VALAXECU5,QBU_VALAXECU6,QBU_VALAXECU7,QBU_VALAXECU8,'+
             'QBU_VALAXECU9,QBU_VALAXECU10,QBU_LIBVALAXECU1,QBU_LIBVALAXECU2,'+
             'QBU_LIBVALAXECU3,QBU_LIBVALAXECU4,QBU_LIBVALAXECU5,'+
             'QBU_LIBVALAXECU6,QBU_LIBVALAXECU7,QBU_LIBVALAXECU8,'+
             'QBU_LIBVALAXECU9,QBU_LIBVALAXECU10,'+
             'SUM(QBU_CAVH),SUM(QBU_CAP),SUM(QBU_CAV),SUM(QBU_CAN),SUM(QBU_CAH),'+  
             'SUM(QBU_QTEVH),SUM(QBU_QTEP),SUM(QBU_QTEV),SUM(QBU_QTEN),SUM(QBU_QTEH),'+
             'SUM(QBU_NBCLTV),SUM(QBU_NBCLTP),SUM(QBU_NBCLTH),SUM(QBU_NBCLTN) '+
             'FROM QBPCALCULTMP WHERE QBU_CODESESSION="'+codeSession+'" GROUP BY '+
             'QBU_VALAXECU1,QBU_VALAXECU2,QBU_VALAXECU3,QBU_VALAXECU4,'+
             'QBU_VALAXECU5,QBU_VALAXECU6,QBU_VALAXECU7,QBU_VALAXECU8,'+
             'QBU_VALAXECU9,QBU_VALAXECU10,QBU_LIBVALAXECU1,QBU_LIBVALAXECU2,'+
             'QBU_LIBVALAXECU3,QBU_LIBVALAXECU4,QBU_LIBVALAXECU5,'+
             'QBU_LIBVALAXECU6,QBU_LIBVALAXECU7,QBU_LIBVALAXECU8,'+
             'QBU_LIBVALAXECU9,QBU_LIBVALAXECU10','BPCoeff (CalculCoeff).',true);
 while not Q.eof do
  begin
   i:=i+1;
   axe1 := Q.fields[0].asString;
   axe2 := Q.fields[1].asString;
   axe3 := Q.fields[2].asString;
   axe4 := Q.fields[3].asString;
   axe5 := Q.fields[4].asString;
   axe6 := Q.fields[5].asString;
   axe7 := Q.fields[6].asString;
   axe8 := Q.fields[7].asString;
   axe9 := Q.fields[8].asString;
   axe10 := Q.fields[9].asString;
   lib1 := Q.fields[10].asString;
   lib2 := Q.fields[11].asString;
   lib3 := Q.fields[12].asString;
   lib4 := Q.fields[13].asString;
   lib5 := Q.fields[14].asString;
   lib6 := Q.fields[15].asString;
   lib7 := Q.fields[16].asString;
   lib8 := Q.fields[17].asString;
   lib9 := Q.fields[18].asString;
   lib10 := Q.fields[19].asString;

   lib1 := '';
   lib2 := '';
   lib3 := '';
   lib4 := '';
   lib5 := '';
   lib6 := '';
   lib7 := '';
   lib8 := '';
   lib9 := '';
   lib10 := '';

   cahistovu := Q.fields[20].asFloat;
   cahistononvu := Q.fields[21].asFloat;
   careelvu := Q.fields[22].asFloat;
   careelnonvu := Q.fields[23].asFloat;
   cahisto := Q.fields[24].asFloat;

   if CalculCoeffQte
    then
     begin
      cahistovu := Q.fields[25].asFloat;
      cahistononvu := Q.fields[26].asFloat;
      careelvu := Q.fields[27].asFloat;
      careelnonvu := Q.fields[28].asFloat;
      cahisto := Q.fields[29].asFloat;
     end;
  { if BPOkOrli
    then
     begin
      if not ChercheSessionObjectif(codeSession,SessionObjectif)
       then cahisto := Q.fields[24].asFloat+cahistovu //cahistononvu+
       else cahisto := Q.fields[24].asFloat
     end
    else cahisto := Q.fields[24].asFloat+cahistononvu+cahistovu;  }
   nbCltVu := Q.fields[30].asInteger;
   nbCltHistoNonVu := Q.fields[31].asInteger;
   nbCltHisto := Q.fields[32].asInteger;
   nbCltHistoReelNonVu := Q.fields[33].asInteger;
   CAProspect := 0;
   extra :='1';
   CAReelVux:=careelvu-careelnonvu;
   if CAReelVux<0
    then CAReelVux:=0;
   if cahistovu<>0
    then Xvisit:=CAReelVux/cahistovu//careelvu/cahistovu
    else Xvisit:=1;
   if CAProspect>careelnonvu
    then NewProspect:=CAProspect
    else NewProspect:=careelnonvu;
   ResteAVisiter:=cahisto-CAProspect-cahistovu-cahistononvu;
   //compare le % du ca histo vu par rapport au (ca histo - ca prospect)
   //avec le % seuil definit en paramètre
   //si < alors XVisit=1
   PrctCaHVu:=100;
   if cahisto-caprospect<>0
    then PrctCaHVu:=(cahistovu/(cahisto-caprospect)*100);
   if PrctCaHVu<PrctSeuilCaHVu
    then Xvisit:=1;

   caretenu:=(ResteAVisiter*Xvisit)+careelvu+NewProspect;

//   if coeff<>1
//    then
//     begin
      //on prend que les lignes qui sont extrapolables
      if (careelvu)<>0
       then coeff:=caretenu/careelvu//coeff:=caretenu/(careelvu+careelnonvu)
       else coeff:=0;

      MExecuteSql('INSERT INTO QBPCOEFF (QBT_CODESESSION,QBT_NUMCOEFF,'+
               'QBT_CODEAXEC1,QBT_CODEAXEC2,QBT_CODEAXEC3,QBT_CODEAXEC4,'+
               'QBT_CODEAXEC5,QBT_CODEAXEC6,QBT_CODEAXEC7,QBT_CODEAXEC8,'+
               'QBT_CODEAXEC9,QBT_CODEAXEC10,'+
               'QBT_LIBC1,QBT_LIBC2,QBT_LIBC3,QBT_LIBC4,'+
               'QBT_LIBC5,QBT_LIBC6,QBT_LIBC7,QBT_LIBC8,'+
               'QBT_LIBC9,QBT_LIBC10,QBT_CAHISTOVU,QBT_CAHISTONONVU,'+
               'QBT_CAREELVU,QBT_CAREELNEW,'+
               'QBT_COEFFEXTRAPO,QBT_CARETENU,QBT_COEFFRETENU,QBT_CAHISTO,'+
               'QBT_NBCLTHISTO,QBT_NBCLTVU,QBT_NBCLTHISTONVU,'+
               'QBT_NBCLTREELNONVU,QBT_CAPROSPECT,QBT_XVISIT,'+
               'QBT_NEWPROSPECT,QBT_RESTEAVISITER,QBT_EXTRAPOLABLE) ' +
               'VALUES ("'+codeSession+'","'+strFPoint4(i)+'","'+axe1+'","'+axe2+
               '","'+axe3+'","'+axe4+'","'+axe5+'","'+axe6+'","'+
               axe7+'","'+axe8+'","'+axe9+'","'+axe10+'","'+
               lib1+'","'+lib2+
               '","'+lib3+'","'+lib4+'","'+lib5+'","'+lib6+'","'+
               lib7+'","'+lib8+'","'+lib9+'","'+lib10+'","'+
               strFPoint4(cahistovu)+
               '","'+strFPoint4(cahistononvu)+'","'+strFPoint4(careelvu)+
               '","'+strFPoint4(careelnonvu)+
               '","'+strFPoint4(coeff)+'","'+strFPoint4(caretenu)+
               '","'+strFPoint4(coeff)+'","'+strFPoint4(cahisto)+
               '","'+strFPoint4(nbCltHisto)+'","'+strFPoint4(nbCltVu)+
               '","'+strFPoint4(nbCltHistoNonVu)+
               '","'+strFPoint4(nbCltHistoReelNonVu)+
               '","'+strFPoint4(CAProspect)+
               '","'+strFPoint4(Xvisit)+
               '","'+strFPoint4(NewProspect)+
               '","'+strFPoint4(ResteAVisiter)+'","'+extra+'")',
               'BPCoeff (RemplitTableQBP).');
    // end;
   Q.next; 
  end;
  ferme(Q);
  MAJExtrapolable(codesession);

  if DonneParamB(pb_BPCoeffGene)
   then MiseAJourDesCoeffsAvecCoeffGene(codeSession);

end;


end.
