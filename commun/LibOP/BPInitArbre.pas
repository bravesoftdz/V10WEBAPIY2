unit BPInitArbre;

interface

uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$else}
     uTob,
{$ENDIF}
     SysUtils,Hctrls,Hent1,Classes, ed_tools;

procedure InitArbre(const codeSession,BPInitialise:hString;coeff1,UpdateSession:boolean);
//Boolean pour la récupération des anciennes sessions
var SQLError, MiseANiveau:boolean;

implementation

uses uutil, UQTobIns, CstCommun, BPFctArbre, BPFctSession,
     BPBasic, BPMaille, BPCoeff, BPEclatement, SynScriptBP,BPUtil,StrUtils,HMsgBox;


procedure MakeArbre(const codeSession:hString);
var Niveau,i,j,NumNoeud,NumNoeudDetail,NumNoeudRef:integer;
    codeGroupBy,CodeValeurAxe:hString;
    TabCodeAxe,TabNumValAxe:array[1..11] of hString;
    TabAxe:array [0..15] of hString;
    Q:TQuery;
    ValeurAxe,Devise,LibValeurAxe:hString;
    QteRef,Ref1,Ref2,Ref3,Ref4,Ref5,Ref6:double;        
    Coeff,CoeffRetenu,Histo,Prevu,Realise,TotalRef,VueRef,Perdu:double;
    ResteAVisiter,TotalCourant,VuCourant,Nouveau,Prospect,NewProspect:double;
    Xvisit,CoeffGene,NbCltRef,NbCltPerdu,NbCltNouveau,NbCltVu,NbCltRestAVoir:double;
    codeWhere:hString;
    QteRetenue,HistoCA,PrevuCA,RealiseCA,CAretenu:double;
    QTIns:TQTobIns;
    DatePiece:TDateTime;
begin
 Niveau:=ChercheNivMax(codeSession);
 if SessionDelai(codeSession)
  then Niveau:=Niveau+1;
 NumNoeud:=BPIncrementeNumNoeud(codeSession);
 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 for i:=1 to 15 do
  TabAxe[i]:='';

 Coeff:=0;
 CoeffRetenu:=0;
 TotalRef:=0;
 VueRef:=0;
 Perdu:=0;
 ResteAVisiter:=0;
 TotalCourant:=0;
 VuCourant:=0;
 Nouveau:=0;
 Prospect:=0;
 NewProspect:=0;
 Xvisit:=0;
 CoeffGene:=0;
 NbCltRef:=0;
 NbCltPerdu:=0;
 NbCltNouveau:=0;
 NbCltVu:=0;
 NbCltRestAVoir:=0;

 //insertion des branches

 for i:=Niveau downto 2 do
  begin
   codeGroupBy:='';
   for j:=1 to i-1 do
    codeGroupBy:=codeGroupBy+',QBR_VALAXENIV'+IntToStr(j);
   CodeValeurAxe:='QBR_VALAXENIV'+IntToStr(i);

   if MiseANiveau=false then begin
   //création de la tob
   QTIns:=TQTobIns.create('QBPARBRE',
                        ['CODESESSION','NUMNOEUD','NUMNOEUDPERE',
                         'CODEAXE','VALEURAXE','LIBVALAXE',
                         'NIVEAU','DEVISE','COMMENTAIREBP',
                         'DATEDELAI','VALBLOQUE','VALBLOQUETMP',
                         'VALAXENIV1','VALAXENIV2','VALAXENIV3','VALAXENIV4',
                         'VALAXENIV5','VALAXENIV6','VALAXENIV7','VALAXENIV8',
                         'VALAXENIV9','VALAXENIV10','VALAXENIV11','VALAXENIV12',
                         'VALAXENIV13','VALAXENIV14','VALAXENIV15',
                         'QTEREF','QTEREFPRCT','QTEC','QTECPRCT',
                         'REF1','REFPRCT1','OP1','OPPRCT1',
                         'REF2','REFPRCT2','OP2','OPPRCT2',
                         'REF3','REFPRCT3','OP3','OPPRCT3',
                         'REF4','REFPRCT4','OP4','OPPRCT4',
                         'REF5','REFPRCT5','OP5','OPPRCT5',
                         'REF6','REFPRCT6','OP6','OPPRCT6',
                         'COEFFCALCUL','COEFFRETENU','HISTO','PREVU','REALISE',
                         'TOTALREF','VUREF','PERDU','RESTEAVISTER',
                         'TOTALCOURANT','VUCOURANT','NOUVEAU',
                         'PROSPECT','NEWPROSPECT','XVISIT','COEFFGENE',
                         'NBCLTREF','NBCLTPERDU','NBCLTNOUVEAU',
                         'NBCLTVU','NBCLTRESTAVOIR','QTERETENUE',
                         'HISTOCA','PREVUCA','REALISECA','CARETENU']);

   Q:=MOpenSql('SELECT QBR_DEVISE'+
               codeGroupBy+
               ',SUM(QBR_QTEREF),'+
               'SUM(QBR_REF1),'+
               'SUM(QBR_REF2),'+
               'SUM(QBR_REF3),'+
               'SUM(QBR_REF4),'+
               'SUM(QBR_REF5),'+
               'SUM(QBR_REF6),'+
               'SUM(QBR_REALISE),'+
               'SUM(QBR_PREVU),'+
               'SUM(QBR_HISTO),'+
               'SUM(QBR_QTERETENUE), '+
               'SUM(QBR_REALISECA),'+
               'SUM(QBR_PREVUCA),'+
               'SUM(QBR_HISTOCA), '+
               'SUM(QBR_CARETENU) '+
               'FROM QBPARBRE '+
               'WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NIVEAU="'+STRFPOINT(i)+
               '" GROUP BY QBR_DEVISE'+codeGroupBy,
               'MakeArbre (BPInitArbre).',true);

   while not Q.eof do
    begin
     NumNoeud:=NumNoeud+1;
     LibValeurAxe:='';
     Devise:=Q.fields[0].asString;
     for j:=1 to i-1 do
      TabAxe[j]:=Q.fields[j].asString;
     QteRef:=Q.fields[i].asFloat;
     Ref1:=Q.fields[i+1].asFloat;
     Ref2:=Q.fields[i+2].asFloat;
     Ref3:=Q.fields[i+3].asFloat;
     Ref4:=Q.fields[i+4].asFloat;
     Ref5:=Q.fields[i+5].asFloat;
     Ref6:=Q.fields[i+6].asFloat;
     Realise:=Q.fields[i+7].asFloat;
     Prevu:=Q.fields[i+8].asFloat;
     Histo:=Q.fields[i+9].asFloat;
     QteRetenue:=Q.fields[i+10].asFloat;
     RealiseCA:=Q.fields[i+11].asFloat;
     PrevuCA:=Q.fields[i+12].asFloat;
     HistoCA:=Q.fields[i+13].asFloat;       
     CAretenu:=Q.fields[i+14].asFloat;

     ValeurAxe:=TabAxe[i-1];
     TabAxe[i-1]:='';

     { EVI / iDate1900 pour QBR_DATEDELAI (génère 31/12/1899 avec 0) }
     QTIns.addValeur([codeSession,numNoeud,'0',TabCodeAxe[i],ValeurAxe,
                    LibValeurAxe,i-1,Devise,'',
                    iDate1900,'-','-',
                    TabAxe[1],TabAxe[2],TabAxe[3],TabAxe[4],
                    TabAxe[5],TabAxe[6],TabAxe[7],TabAxe[8],
                    TabAxe[9],TabAxe[10],TabAxe[11],TabAxe[12],
                    TabAxe[13],TabAxe[14],TabAxe[15],
                    QteRef,0,QteRef,0,Ref1,0,Ref1,0,Ref2,0,Ref2,0,Ref3,0,Ref3,0,Ref4,0,Ref4,0,
                    Ref5,0,Ref5,0,Ref6,0,Ref6,0,Coeff,CoeffRetenu,Histo,Prevu,
                    Realise,TotalRef,VueRef,Perdu,ResteAVisiter,TotalCourant,
                    VuCourant,Nouveau,Prospect,NewProspect,Xvisit,CoeffGene,NbCltRef,
                    NbCltPerdu,NbCltNouveau,NbCltVu,NbCltRestAVoir,QteRetenue,
                    HistoCA,PrevuCA,RealiseCA,CAretenu]);


  {   MExecuteSql('INSERT INTO QBPARBRE '+
               '(QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_CODEAXE,QBR_VALEURAXE,'+
               'QBR_LIBVALAXE,QBR_NIVEAU,QBR_DEVISE,QBR_COMMENTAIREBP,'+
               'QBR_DATEDELAI,QBR_VALBLOQUE,QBR_VALBLOQUETMP,'+
               'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,'+
               'QBR_VALAXENIV5,QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,'+
               'QBR_VALAXENIV9,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,'+
               'QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,'+
               'QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_QTECPRCT,'+
               'QBR_REF1,QBR_REFPRCT1,QBR_OP1,QBR_OPPRCT1,'+
               'QBR_REF2,QBR_REFPRCT2,QBR_OP2,QBR_OPPRCT2,'+
               'QBR_REF3,QBR_REFPRCT3,QBR_OP3,QBR_OPPRCT3,'+
               'QBR_REF4,QBR_REFPRCT4,QBR_OP4,QBR_OPPRCT4,'+
               'QBR_REF5,QBR_REFPRCT5,QBR_OP5,QBR_OPPRCT5,'+
               'QBR_REF6,QBR_REFPRCT6,QBR_OP6,QBR_OPPRCT6,'+
               'QBR_COEFFCALCUL,QBR_COEFFRETENU,QBR_HISTO,QBR_PREVU,QBR_REALISE,'+
               'QBR_TOTALREF,QBR_VUREF,QBR_PERDU,QBR_RESTEAVISTER,'+
               'QBR_TOTALCOURANT,QBR_VUCOURANT,QBR_NOUVEAU,'+
               'QBR_PROSPECT,QBR_NEWPROSPECT,QBR_XVISIT,QBR_COEFFGENE,'+
               'QBR_NBCLTREF,QBR_NBCLTPERDU,QBR_NBCLTNOUVEAU,'+
               'QBR_NBCLTVU,QBR_NBCLTRESTAVOIR,QBR_QTERETENUE) '+
               'VALUES ("'+codeSession+'","'+STRFPOINT(NumNoeud)+'","0","'+
               TabCodeAxe[i]+'","'+ValeurAxe+'","'+LibValeurAxe+'","'+STRFPOINT(i-1)+
               '","'+Devise+'","","'+USDATETIME(0)+'","-","-","'+
               TabAxe[1]+'","'+TabAxe[2]+'","'+TabAxe[3]+
               '","'+TabAxe[4]+'","'+TabAxe[5]+'","'+TabAxe[6]+'","'+TabAxe[7]+
               '","'+TabAxe[8]+'","'+TabAxe[9]+'","'+TabAxe[10]+'","'+TabAxe[11]+
               '","'+TabAxe[12]+'","'+TabAxe[13]+'","'+TabAxe[14]+'","'+TabAxe[15]+
               '","'+StrFPoint(QteRef)+'","'+StrFPoint(0)+'","'+StrFPoint(QteRef)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref1)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref1)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref2)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref2)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref3)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref3)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref4)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref4)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref5)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref5)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Ref6)+'","'+StrFPoint(0)+'","'+StrFPoint(Ref6)+'","'+StrFPoint(0)+
               '","'+StrFPoint(Coeff)+'","'+StrFPoint(CoeffRetenu)+
               '","'+StrFPoint(Histo)+'","'+StrFPoint(Prevu)+'","'+StrFPoint(Realise)+
               '","'+StrFPoint(TotalRef)+'","'+StrFPoint(VueRef)+
               '","'+StrFPoint(Perdu)+'","'+StrFPoint(ResteAVisiter)+
               '","'+StrFPoint(TotalCourant)+'","'+StrFPoint(VuCourant)+'","'+StrFPoint(Nouveau)+
               '","'+StrFPoint(Prospect)+'","'+StrFPoint(NewProspect)+
               '","'+StrFPoint(Xvisit)+'","'+StrFPoint(CoeffGene)+
               '","'+StrFPoint(NbCltRef)+'","'+StrFPoint(NbCltPerdu)+'","'+StrFPoint(NbCltNouveau)+
               '","'+StrFPoint(NbCltVu)+'","'+StrFPoint(NbCltRestAVoir)+
               '","'+StrFPoint(QteRetenue)+
               '")',
               'MakeArbre (BPInitArbre).');     }
     Q.next;
     end;
   ferme(Q);
   QTIns.free;
  end;
 //Mise à jour des noeuds peres
 MiseAJourNoeudPere(codeSession);
 //Mise a jour des % histo
 MiseAJourPrctArbre(codeSession,true,true);
 //Mise à jour des % de variation
 MAJPrctVariation(codeSession);
   end;
end;




function RechSqlInsertFeuillesObjCasOBJ(const codeSession,SessionObjectif: hString; niveau:integer): hString;
var codeRestriction,CodeJoinR,codeChpsSum,codeWhere,chpsLib,codeTable:hString;
    i : integer;
begin
  //cherche code restriction de la session
  codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);
  codeChpsSum:=' QBR_QTEC,"0",QBR_OP1,QBR_OP2,QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6';
  codeTable:=' QBPARBRE ';
  codeWhere:=' QBR_CODESESSION="'+SessionObjectif+'" ';
  for i:=1 to Niveau-1 do
     ChpsLib:=ChpsLib+',QBR_VALAXENIV'+IntToStr(i);
  ChpsLib:=ChpsLib+',QBR_VALEURAXE,QBR_LIBVALAXE,QBR_DEVISE';
  result:='SELECT '+codeChpsSum+ChpsLib+
           ' FROM '+codeTable+
           ' WHERE '+codeWhere+' '+codeRestriction+
           ' AND QBR_NIVEAU="'+STRFPOINT(niveau)+'"';
end;

//-----------------> ORLI
function RechSqlInsertFeuillesObjCasORLI(okPremMaille,okDerMaille:boolean;
         const codeSession,SaiCmd,SaiCmdRef:hString;niveau:integer;
         DateRefDeb,DateRefFin,DateCDeb,DateCFin:TDateTime;
         const TabNumValAxe:array of hString):hString;
var codeRestriction,codeChps,codeChps2,codeChps3,codeChpsSum,codeTable,codeWhere:hString;
    ChpsLib23,codeChpsGroupBy:hString;
    codeWherePeriodeRef,codeWherePeriodeC,codeWherePeriodeRefP,structure:hString;
    ChpsLib: hString;
    CodeJoinR: hString;
    i:integer;
begin
  Result:='';
  //cherche code restriction de la session
  codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);
  codeChps:='';
  codeChps2:='';
  codeChps3:='';
  structure:=ChercheSessionStructure(codeSession);
  //cherche le code sql champs
  for i:=1 to Niveau do
   begin
    if codeChps2=''
     then codeChps:=' QBV_VALAXEP'+TabNumValAxe[i]
     else codeChps:=codeChps+',QBV_VALAXEP'+TabNumValAxe[i];
    if codeChps2=''
     then codeChps2:=' QBV_VALAXEP'+TabNumValAxe[i]
     else codeChps2:=codeChps2+'||QBV_VALAXEP'+TabNumValAxe[i];
    if codeChps3=''
     then codeChps3:=' P.QBV_VALAXEP'+TabNumValAxe[i]
     else codeChps3:=codeChps3+'||P.QBV_VALAXEP'+TabNumValAxe[i];
   end;
  ChpsLib:=',"L"';
  ChpsLib23:='||"L"';
  codeChpsGroupBy:=codeChps+',QBV_DEVISE ';
  codeChps:=codeChps+ChpsLib+',QBV_DEVISE ';
  codeChps2:=codeChps2+ChpsLib23+'||QBV_DEVISE ';
  codeChps3:=codeChps3+ChpsLib23+'||P.QBV_DEVISE ';

  //cherche le code sql champs
  codeChpsSum:=' SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
               'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
               'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
               'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
               'SUM(QBV_QTELT1+QBV_QTELT2+QBV_QTELT3+QBV_QTELT4+QBV_QTELT5+'+
               'QBV_QTELT6+QBV_QTELT7+QBV_QTELT8+QBV_QTELT9+QBV_QTELT10+'+
               'QBV_QTELT11+QBV_QTELT12+QBV_QTELT13+QBV_QTELT14+QBV_QTELT15+'+
               'QBV_QTELT16+QBV_QTELT17+QBV_QTELT18+QBV_QTELT19+QBV_QTELT20),'+
               'SUM(QBV_QTE1),SUM(QBV_QTE2),SUM(QBV_QTE3),'+
               'SUM(QBV_QTE4),SUM(QBV_QTE5),0 ';

  //cherche le code sql table
  codeTable:=' QBPPIVOT ';

  if okPremMaille and (SaiCmdRef<>'')
   then
    begin
     codeWherePeriodeRef:=' AND QBV_DATEPIVOT<="'+USDATETIME(DateRefFin)+'" ';
     codeWherePeriodeRefP:=' AND  P.QBV_DATEPIVOT<="'+USDATETIME(DateRefFin)+'" ';
     codeWherePeriodeC:=' AND QBV_DATEPIVOT<="'+USDATETIME(DateCFin)+'" ';
    end else
   if okDerMaille and (SaiCmdRef<>'')
    then
     begin
      codeWherePeriodeRef:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateRefDeb)+'" ';
      codeWherePeriodeRefP:=' AND P.QBV_DATEPIVOT>="'+USDATETIME(DateRefDeb)+'" ';
      codeWherePeriodeC:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateCDeb)+'" ';
     end
      else
       begin
         //cherche le code sql where  période
        codeWherePeriodeRef:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateRefDeb)+
                             '" AND QBV_DATEPIVOT<="'+USDATETIME(DateRefFin)+'" ';
        codeWherePeriodeRefP:=' AND P.QBV_DATEPIVOT>="'+USDATETIME(DateRefDeb)+
                              '" AND P.QBV_DATEPIVOT<="'+USDATETIME(DateRefFin)+'" ';
        codeWherePeriodeC:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateCDeb)+
                           '" AND QBV_DATEPIVOT<="'+USDATETIME(DateCFin)+'" ';
      end;

  if (SaiCmdRef<>'') //and (not okDelai)
   then
    begin
     codeWherePeriodeRef:=codeWherePeriodeRef+' AND QBV_VALAXEP8="'+SaiCmdRef+'" ';
     codeWherePeriodeRefP:=codeWherePeriodeRefP+' AND P.QBV_VALAXEP8="'+SaiCmdRef+'" ';
    end;

  if (SaiCmd<>'') //and (not okDelai)
   then codeWherePeriodeC:=codeWherePeriodeC+' AND QBV_VALAXEP8="'+SaiCmd+'" ';

  //cherche le code sql where
  codeWhere:=' QBV_CODESTRUCT="'+structure+'" '+codeWherePeriodeRef+codeRestriction;

  result:='SELECT '+codeChpsSum+','+codeChps+' FROM '+codeTable+
          ' WHERE '+codeWhere+' GROUP BY '+codeChpsGroupBy+
          ' UNION SELECT 0,0,0,0,0,0,0,0,'+codeChps+' FROM '+codeTable+
          ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWherePeriodeC+' '+codeRestriction+
          ' AND '+codeChps2+' NOT IN (SELECT DISTINCT '+codeChps3+
          ' FROM '+codeTable+' P WHERE P.QBV_CODESTRUCT="'+structure+'" AND '+
          codeWherePeriodeRefP+')';
end;
//ORLI <-----------------

function RechSqlInsertFeuillesObjCasPGI(const codeSession,BPInit:String;niveau:integer;
         DateRefDeb,DateRefFin,DateCDeb:TDateTime; const TabCodeAxe:array of hString):hString;
var codeRestriction,codeChps,codeChps2,codeChps3,codeChpsSum,codeChpsSumDet,codeTable,codeWhere,notuse:String;
    ChpsLibGB:hString;
    ChpsLib,joinLib: hString;
    CodeJoinR: hString;
    j:integer;
    TS, TS2, TS3 : TstringList;
    s, LJSQL, Alias, AliasOk, AliasSup : String;
    NbMaille,NbValAff:integer;

        function NbFinMoisSession (DateDeb,DateFin : TDateTime) : string;
        Var DateTMP : TDateTime;
        begin
          DateTMP:=FinDeMois(DateDeb);
          While DateTmp <= DateFin do
          begin
            if result = '' then result:='PPU_DATEFIN="'+USDATETIME(DateTMP) +'"'
            else result:=result + ' OR PPU_DATEFIN="'+USDATETIME(DateTMP) +'"';
            DateTMP:=DateTMP+1;
            DateTMP:=FinDeMois(DateTMP);
          end;
        end;

begin
  result:='';
  //cherche code restriction de la session
  codeRestriction:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,CodeJoinR);

  codeChps:='';
  codeChps2:='';
  codeChps3:='';
  joinLib:=codejoinR;
  DonneCodeChpJoinPGINbNivLib(TabCodeAxe,niveau,ChpsLib,ChpsLibGB,joinLib,Codesession);

  Case ContextBP of
    0,1 : begin //Mode-GC
          //cherche le code sql champs
          if AxeDeviseOK(codeSession) then
          codeChpsSum:=' SUM(GL_QTEFACT),0,SUM(GL_TOTALTTCDEV),SUM(GL_TOTALHTDEV),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PUHTNETDEV*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNETDEV*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(GL_TOTALHTDEV)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)), '
          else
          codeChpsSum:=' SUM(GL_QTEFACT),0,SUM(GL_TOTALTTC),SUM(GL_TOTALHT),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
                       'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)), ';

          //cherche le code sql table
          codeTable:=' LIGNE ';
          //cherche le code sql where
          // MODIF LM 08/02/2006 pour SPAG
          // Ajouter une clause pour ne prendre en compte que les articles marchandises+prestations
          codeWhere:=' GL_DATEPIECE>="'+USDATETIME(DateRefDeb)+
                     '" AND GL_DATEPIECE<="'+USDATETIME(DateRefFin)+
                     '" AND GL_TYPELIGNE="ART" AND GL_TYPEARTICLE IN ("MAR", "PRE", "NOM") ';
        end;
    2 : begin //Compta
          //cherche le code sql champs
          if AxeDeviseOK(codeSession) then
          codeChpsSum:=' 0,0,SUM(Y_DEBITDEV-Y_CREDITDEV),SUM(Y_CREDITDEV-Y_DEBITDEV),0,0,0,0, '
          else
          codeChpsSum:=' 0,0,SUM(Y_DEBIT-Y_CREDIT),SUM(Y_CREDIT-Y_DEBIT),0,0,0,0, ';
          //cherche le code sql table
          codeTable:=' ANALYTIQ ';
          //cherche le code sql where
          codeWhere:=' Y_DATECOMPTABLE>="'+USDATETIME(DateRefDeb)+
                     '" AND Y_DATECOMPTABLE<="'+USDATETIME(DateRefFin)+'" ';
        end;

    3 : begin //Paie
          ReqValAff(codeSession,ChpsLib,niveau,NbValAff,codeChpsSum,codeChpsSumDet,codeTable,codeWhere,notuse,notuse,notuse,notuse);

          codeWhere := AnsiReplaceText(codeWhere,'[DATEDEBUT]',USDATETIME(DateRefDeb));
          codeWhere := AnsiReplaceText(codeWhere,'[DATEFIN]',USDATETIME(DateRefFin));
          if DonneMethodeSession(codeSession)='2' then
          begin
            codeChpsSum := AnsiReplaceText(codeChpsSum,'[EFFECTIFFINMOIS]',NbFinMoisSession(DateRefDeb,DateRefFin));
            codeChpsSumDet := AnsiReplaceText(codeChpsSumDet,'[EFFECTIFFINMOIS]',NbFinMoisSession(DateRefDeb,DateRefFin));
          end
          else
          begin
            codeChpsSum := AnsiReplaceText(codeChpsSum,'[EFFECTIFFINMOIS]','PPU_DATEFIN="[DATEFIN]"');
            codeChpsSumDet := AnsiReplaceText(codeChpsSumDet,'[EFFECTIFFINMOIS]','PPU_DATEFIN="[DATEFIN]"');
          end;
          codeChpsSum := AnsiReplaceText(codeChpsSum,'[DATEDEBUT]',USDATETIME(DateRefDeb));
          codeChpsSum := AnsiReplaceText(codeChpsSum,'[DATEFIN]',USDATETIME(DateRefFin));
          codeChpsSumDet := AnsiReplaceText(codeChpsSumDet,'[DATEDEBUT]',USDATETIME(DateRefDeb));
          codeChpsSumDet := AnsiReplaceText(codeChpsSumDet,'[DATEFIN]',USDATETIME(DateRefFin));

          NbMaille:=NbMoisIntervalle(DateRefDeb,DateRefFin);
          //Gestion des périodes
          //if BPInit='0' //Loi d'Eclatement
          if BPInit='2' then NbMaille := NbMaille div 4; //Semaine
          if BPInit='3' then NbMaille := NbMaille div 2;//Quinzaine
          //if BPInit='4' //Mois
          //if BPInit='5' //Mois 4-4-5
          if BPInit='6' then NbMaille := NbMaille * 3;//Trimestre
          if BPInit='7' then NbMaille := NbMaille * 4;//Quadrimestre
          codeChpsSum := AnsiReplaceText(codeChpsSum,'[NBMOIS]',intToStr(NbMaille));
          codeChpsSumDet := AnsiReplaceText(codeChpsSumDet,'[NBMOIS]',intToStr(NbMaille));

        end
   end;//CASE
  LJSQL := '';
  TS := TStringList.Create;
  try
    TS2 := TStringList.Create;
    try
      TS3 := TStringList.Create;
      try
        LJSQL := AnalyseJoinSQL(joinLib, TS, TS2, TS3);
        //Suppression des alias des champs
        for j := 0 to TS2.Count - 1 do
        begin
          s := TS2[j];
          Alias   := ReadTokenSt(s);
          AliasOk := ReadTokenSt(s);
          if not StrToBool(AliasOk) then
          begin
            AliasSup := Alias + ';' + AliasSup;
            if ((TS3[j] <> '') AND (Pos(TS3[j],AliasSup)=0)) then
            begin
              codeChpsSum      := StringReplace(codeChpsSum    , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
              codeChpsSumDet   := StringReplace(codeChpsSumDet , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
              ChpsLib          := StringReplace(ChpsLib        , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
              codeRestriction  := StringReplace(codeRestriction, Alias+'.', TS3[j]+'.', [rfReplaceAll]);
              ChpsLibGB        := StringReplace(ChpsLibGB      , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
              LJSQL            := StringReplace(LJSQL          , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
            end
            else
            begin
              codeChpsSum      := StringReplace(codeChpsSum    , Alias+'.', '', [rfReplaceAll]);
              codeChpsSumDet   := StringReplace(codeChpsSumDet , Alias+'.', '', [rfReplaceAll]);
              ChpsLib          := StringReplace(ChpsLib        , Alias+'.', '', [rfReplaceAll]);
              codeRestriction  := StringReplace(codeRestriction, Alias+'.', '', [rfReplaceAll]);
              ChpsLibGB        := StringReplace(ChpsLibGB      , Alias+'.', '', [rfReplaceAll]);
              LJSQL            := StringReplace(LJSQL          , Alias+'.', '', [rfReplaceAll]);
            end;
          end;
        end;
      finally
        if Assigned(TS3) then TS3.Free;
      end;
    finally
      if Assigned(TS2) then TS2.Free;
    end;
  finally
    if Assigned(TS) then TS.Free;
  end;

  ChpsLibGB        := StringReplace(ChpsLibGB      , 'TRIM', '', [rfReplaceAll]);
  ChpsLibGB        := StringReplace(ChpsLibGB      , '(', '', [rfReplaceAll]);
  ChpsLibGB        := StringReplace(ChpsLibGB      , ')', '', [rfReplaceAll]);

  //Vérifie si la table SALARIE est principale et l'axe SALARIE est le dernier
  if ContextBP=3 then LJSQL := AnalyseJoinSQLPaie(LJSQL,codeTable);

  result:='SELECT '+codeChpsSum+ChpsLib+'[CHAMPDETAIL] FROM '+codeTable+' '+LJSQL+' '+
          ' WHERE '+codeWhere+' '+codeRestriction+' GROUP BY '+ChpsLibGB+'[CHAMPDETAILGROUPBY]';
  result:=result+'@||@'+'SELECT '+codeChpsSumDet+ChpsLib+'[CHAMPDETAIL] FROM '+codeTable+' '+LJSQL+' '+
          ' WHERE '+codeWhere+' '+codeRestriction+' GROUP BY '+ChpsLibGB+'[CHAMPDETAILGROUPBY]';

end;

function RechercheCodeSqlInsertFeuillesObj(okPremMaille,okDerMaille:boolean;
         const codeSession,BPInit,SaiCmd,SaiCmdRef:hString;niveau:integer;
         DateRefDeb,DateRefFin,DateCDeb,DateCFin:TDateTime;
         const TabCodeAxe,TabNumValAxe:array of hString):hString;
var SessionObjectif: hString;
begin
  Result:='';
  if ChercheSessionObjectif(codeSession,SessionObjectif) then
     Result:=RechSqlInsertFeuillesObjCasOBJ(codeSession,SessionObjectif,niveau)
            else
   if BPOkOrli then
       //-----------------> ORLI
       Result:=RechSqlInsertFeuillesObjCasOrli(okPremMaille,okDerMaille,
                       codeSession,SaiCmd,SaiCmdRef,niveau,
                       DateRefDeb,DateRefFin,DateCDeb,DateCFin,TabNumValAxe)
       //ORLI <-----------------
             else
          Result:=RechSqlInsertFeuillesObjCasPGI(codeSession,BPInit,niveau,
                         DateRefDeb,DateRefFin,DateCDeb,TabCodeAxe);

end;

procedure InsertFeuillesMailles(okPremMaille,okDerMaille:boolean;
          const codeSession,BpInitialise:hString;
          OkDelai,coeff1:boolean;
          Niveau:integer;
          SaiCmd,SaiCmdRef:hString;
          DateRefDeb,DateRefFin,DateCDeb,DateCFin:TDateTime;
          TabCodeAxe,TabNumValAxe:array of hString;
          var numNoeud,NumNoeudDetail,NumNoeudRef:integer);
var i,niveauX:integer;
    ReqSQL,codeSql0,codeSql,codeSqldetail0,codeSqldetail,LibValAxe,Devise,ValeurAxe:hString;
    Q,QDet,Q2:TQuery;
    QteC,CA1,CA2,CA3,CA4,CA5,CA6:double;
    TabAxe:array [0..15] of hString;
    DateDelai,DatePiece:TDateTime;
    TotalPerRefCA,CCltVuCA,RefCltVuCA,CltPerduCA,NouvCltCA:double;
    TotalPerRefQte,CCltVuQte,RefCltVuQte,CltPerduQte,NouvCltQte:double;
    CltProspectCA,CltProspectQte:double;
    Coeff,CoeffRetenu,Prevu,Realise,TotalRef,vueC:double;
    ResteAVisiter,TotalCourant,NewProspect:double;
    Xvisit,CoeffGene,NbCltRef,NbCltPerdu,NbCltNouveau,NbCltVu,NbCltRestAVoir:double;
    NbCltProspect,NbCltC:double;
    OkSessionObj:boolean;
    QTIns,QTDetIns:TQTobIns;
    Qterealise,Qteprevu,qteRet:double;
    CAretenu:double;
    Extrapo:hString;
    AxeOK:boolean;
    PrefixH,TypeVal:string;
begin
 //initialisation
  for i:=0 to 15 do TabAxe[i]:='';

 PrefixH:='';
 TypeVal:='';
 AxeOK:=false;
 Coeff:=0;
 CoeffRetenu:=0;
 TotalPerRefCA:=0;
 Prevu:=0;
 Realise:=0;
 TotalRef:=0;
 RefCltVuCA:=0;
 vueC:=0;
 CltPerduCA:=0;
 ResteAVisiter:=0;
 TotalCourant:=0;
 CCltVuCA:=0;
 NouvCltCA:=0;
 NewProspect:=0;
 Xvisit:=0;
 CoeffGene:=0;
 NbCltRef:=0;
 NbCltPerdu:=0;
 NbCltNouveau:=0;
 NbCltVu:=0;
 NbCltRestAVoir:=0;
 TotalPerRefQte:=0;
 CCltVuQte:=0;
 RefCltVuQte:=0;
 CltPerduQte:=0;
 NouvCltQte:=0;
 Qterealise:=0;
 Qteprevu:=0;
 qteRet:=0;
 CAretenu:=0;
 CltProspectCA:=0;
 CltProspectQte:=0;
 NbCltProspect:=0;
 NbCltC:=0;
 Extrapo:='';

 OkSessionObj:=OkSessionObjectif(codeSession);

 //code Sql
 ReqSQL := RechercheCodeSqlInsertFeuillesObj(okPremMaille,okDerMaille,
                 codeSession,BPInitialise,SaiCmd,SaiCmdRef,niveau,
                 DateRefDeb,DateRefFin,DateCDeb,DateCFin,TabCodeAxe,TabNumValAxe);

  Case ContextBP of
    0,1 : begin
          CodeSQL0:=ReadTokenPipe(ReqSQL,'@||@');
          codeSql := AnsiReplaceText(codeSql0,'[PREFIXE]','');
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAIL]','');
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAILGROUPBY]','');
          codeSqldetail := AnsiReplaceText(codeSql0,'[PREFIXE]','');
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAIL]',',GL_DATEPIECE');
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAILGROUPBY]',',GL_DATEPIECE');
        end;
    2 : begin
          CodeSQL0:=ReadTokenPipe(ReqSQL,'@||@');
          codeSql := AnsiReplaceText(codeSql0,'[PREFIXE]','');
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAIL]','');
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAILGROUPBY]','');
          codeSqldetail := AnsiReplaceText(codeSql0,'[PREFIXE]','');
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAIL]',',Y_DATECOMPTABLE');
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAILGROUPBY]',',Y_DATECOMPTABLE');
        end;
    3 : begin
          CodeSQL0:=ReadTokenPipe(ReqSQL,'@||@');
          codeSqldetail0:=ReadTokenPipe(ReqSQL,'@||@');
          DonnePrefixe(codeSession,PrefixH,TypeVal);
          TypeVal := PrefixH;
          if ((PrefixH = 'NBS') OR (PrefixH = 'EFM') OR (PrefixH = 'PSA')) then PrefixH := 'PPU';
          codeSql := AnsiReplaceText(codeSql0,'[PREFIXE]',PrefixH);
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAIL]','');
          codeSql := AnsiReplaceText(codeSql,'[CHAMPDETAILGROUPBY]','');
          codeSqldetail:=AnsiReplaceText(codeSqlDetail0,'[PREFIXE]',PrefixH);
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAIL]',','+PrefixH+'_DATEDEBUT');
          codeSqldetail:=AnsiReplaceText(codeSqldetail,'[CHAMPDETAILGROUPBY]',','+PrefixH+'_DATEDEBUT,'+PrefixH+'_DATEFIN');
        end;
  end;

  //Si il s'agit d'une session en Loi d'Eclatement, on ne remplit pas QBPARBREDETAIL
  if DonneMethodeSession(codeSession) = '2' then codeSqldetail:='';

  if MiseANiveau=false then begin
 //création de la tob
 QTIns:=TQTobIns.create('QBPARBRE',
                        ['CODESESSION','NUMNOEUD','NUMNOEUDPERE',
                         'CODEAXE','VALEURAXE','LIBVALAXE',
                         'NIVEAU','DEVISE','COMMENTAIREBP',
                         'DATEDELAI','VALBLOQUE','VALBLOQUETMP',
                         'VALAXENIV1','VALAXENIV2','VALAXENIV3','VALAXENIV4',
                         'VALAXENIV5','VALAXENIV6','VALAXENIV7','VALAXENIV8',
                         'VALAXENIV9','VALAXENIV10','VALAXENIV11','VALAXENIV12',
                         'VALAXENIV13','VALAXENIV14','VALAXENIV15',
                         'QTEREF','QTEREFPRCT','QTEC','QTECPRCT',
                         'REF1','REFPRCT1','OP1','OPPRCT1',
                         'REF2','REFPRCT2','OP2','OPPRCT2',
                         'REF3','REFPRCT3','OP3','OPPRCT3',
                         'REF4','REFPRCT4','OP4','OPPRCT4',
                         'REF5','REFPRCT5','OP5','OPPRCT5',
                         'REF6','REFPRCT6','OP6','OPPRCT6',
                         'COEFFCALCUL','COEFFRETENU','COEFFGENE',
                         'NBCLTREF','NBCLTPERDU','NBCLTNOUVEAU',
                         'NBCLTVU','NBCLTRESTAVOIR',
                         'NBCLTPROSPECT','NBCLTC',
                         'RESTEAVISTER','XVISIT',
                         'PROSPECT','PROSPECTQTE','NEWPROSPECT',
                         'VUREF','VUCOURANT','NOUVEAU','PERDU',
                         'VUREFQTE','VUCOURANTQTE','NOUVEAUQTE','PERDUQTE',
                         'TOTALREF','TOTALCOURANT',
                         'HISTOCA','REALISECA','PREVUCA','CARETENU',
                         'HISTO','PREVU','REALISE','QTERETENUE',
                         'EXTRAPOLABLE']);
  end;

  QTDetIns:=TQTobIns.create('QBPARBREDETAIL',
                        ['CODESESSION','NUMNOEUD','NIVEAU',
                         'DATEDELAI',
                         'QTEREF','QTEREFPRCT','QTEC','QTECPRCT',
                         'REF1','REFPRCT1','OP1','OPPRCT1',
                         'REF2','REFPRCT2','OP2','OPPRCT2',
                         'REF3','REFPRCT3','OP3','OPPRCT3',
                         'REF4','REFPRCT4','OP4','OPPRCT4',
                         'REF5','REFPRCT5','OP5','OPPRCT5',
                         'REF6','REFPRCT6','OP6','OPPRCT6',
                         'DATEPIECE','NUMNOEUDREF']);

  try
    Q:=OpenSql(codeSql,true);
  except
    on E: Exception do
    begin
      Q := nil;
      SQLError := true;
      exit;
    end;
  end;

  while not Q.eof do
  begin
   numNoeud:=numNoeud+1;

   QteC:=Q.fields[0].asFloat;
   CA1:=Q.fields[2].asFloat;
   CA2:=Q.fields[3].asFloat;
   CA3:=Q.fields[4].asFloat;
   CA4:=Q.fields[5].asFloat;
   CA5:=Q.fields[6].asFloat;
   CA6:=Q.fields[7].asFloat;

    for i:=0 to niveau-1 do TabAxe[i]:=Q.fields[8+i].asString;

   LibValAxe:=Q.fields[8+niveau].asString;
   Devise:=Q.fields[8+niveau+1].asString;

   { EVI / iDate1900 pour DATEDELAI (génère 31/12/1899 avec 0) }
   DateDelai:=iDate1900;

   ValeurAxe:=TabAxe[niveau-1];

   niveauX:=niveau;

   if OkDelai
    then
     begin
      niveauX:=niveau+1;
      TabCodeAxe[niveau]:='DELAI';
      DateDelai:=DateCDeb;
      ValeurAxe:=DateTimeToStr(ChercheNomValeurAxeDelai(BpInitialise,0,DateCDeb));
      LibValAxe:=ValeurAxe;
     end;

   if not OkSessionObj
    then
     begin
      Q2:=MOpenSql('SELECT QBT_COEFFRETENU,QBT_CAHISTO,QBT_QTEHISTO,'+
                   'QBT_CAPROSPECT,QBT_QTEPROSPECT,QBT_NEWPROSPECT,'+
                   'QBT_CAREELVU,QBT_QTEREELVU,QBT_CAHISTOVU,QBT_QTEHISTOVU,'+
                   'QBT_CAHISTONONVU,QBT_QTEHISTONONVU,QBT_CAREELNEW,QBT_QTEREELNEW,'+
                   'QBT_NBCLTHISTO,QBT_NBCLTHISTONVU,QBT_NBCLTREELNONVU,QBT_NBCLTVU,'+
                   'QBT_NBCLTPROSPECT,'+
                   'QBT_RESTEAVISITER,QBT_XVISIT,QBT_COEFFGENE,QBT_EXTRAPOLABLE '+
                   'FROM QBPCOEFF WHERE QBT_CODESESSION="'+codeSession+
                   '" AND QBT_CODEAXEC1="'+TabAxe[0]+
                   '" AND QBT_CODEAXEC2="'+TabAxe[1]+
                   '" AND QBT_CODEAXEC3="'+TabAxe[2]+
                   '" AND QBT_CODEAXEC4="'+TabAxe[3]+
                   '" AND QBT_CODEAXEC5="'+TabAxe[4]+
                   '" AND QBT_CODEAXEC6="'+TabAxe[5]+
                   '" AND QBT_CODEAXEC7="'+TabAxe[6]+
                   '" AND QBT_CODEAXEC8="'+TabAxe[7]+
                   '" AND QBT_CODEAXEC9="'+TabAxe[8]+
                   '" AND QBT_CODEAXEC10="'+TabAxe[9]+
                   '"','BPInitArbre (InsertFeuillesMailles).',true);
      if not Q2.eof
      then
      begin
        coeff:=Q2.fields[0].asFloat;
        if coeff1 then coeff:=1;

        TotalPerRefCA:=Q2.fields[1].asFloat;
        TotalPerRefQte:=Q2.fields[2].asFloat;
        CltProspectCA:=Q2.fields[3].asFloat;
        CltProspectQte:=Q2.fields[4].asFloat;
        NewProspect:=Q2.fields[5].asFloat;
        CCltVuCA:=Q2.fields[6].asFloat;
        CCltVuQte:=Q2.fields[7].asFloat;
        RefCltVuCA:=Q2.fields[8].asFloat;
        RefCltVuQte:=Q2.fields[9].asFloat;
        CltPerduCA:=Q2.fields[10].asFloat;
        CltPerduQte:=Q2.fields[11].asFloat;
        NouvCltCA:=Q2.fields[12].asFloat;
        NouvCltQte:=Q2.fields[13].asFloat;

        NbCltRef:=Q2.fields[14].asFloat;
        NbCltPerdu:=Q2.fields[15].asFloat;
        NbCltNouveau:=Q2.fields[16].asFloat;
        NbCltVu:=Q2.fields[17].asFloat;
        NbCltProspect:=Q2.fields[18].asFloat;
        NbCltRestAVoir:=NbCltRef-NbCltVu-NbCltPerdu;
        NbCltC:=NbCltNouveau+NbCltVu;

        ResteAVisiter:=Q2.fields[19].asFloat;
        Xvisit:=Q2.fields[20].asFloat;
        CoeffGene:=Q2.fields[21].asFloat;

        Extrapo:=Q2.fields[22].asString;

        CoeffRetenu:=coeff;
        Realise:=CCltVuCA+NouvCltCA;
        Qterealise:=CCltVuQte+NouvCltQte;
        Prevu:=realise*coeff;
        Qteprevu:=round(Qterealise*coeff);
        QteRet:=QtePrevu;
        CAretenu:=prevu;

        if DonneParamS(ps_BPCoeffPerCAQte)='2'
        then
        begin
          TotalRef:=TotalPerRefQte;
          TotalCourant:=Qterealise;
        end
        else
        begin
          TotalRef:=TotalPerRefCA;
          TotalCourant:=realise;
        end;
      end;
      ferme(Q2);
    end;

    if MiseANiveau=false then begin
    QTIns.addValeur([codeSession,numNoeud,'0',
                    TabCodeAxe[niveau],ValeurAxe,LibValAxe,
                    niveauX,Devise,'',
                    DateDelai,'-','-',
                    TabAxe[0],TabAxe[1],TabAxe[2],TabAxe[3],
                    TabAxe[4],TabAxe[5],TabAxe[6],TabAxe[7],
                    TabAxe[8],TabAxe[9],TabAxe[10],TabAxe[11],
                    TabAxe[12],TabAxe[13],TabAxe[14],
                    QteC,0,QteC,0,CA1,0,CA1,0,CA2,0,CA2,0,CA3,0,CA3,0,CA4,0,CA4,0,
                    CA5,0,CA5,0,CA6,0,CA6,0,
                    Coeff,CoeffRetenu,CoeffGene,
                    NbCltRef,NbCltPerdu,NbCltNouveau,
                    NbCltVu,NbCltRestAVoir,
                    NbCltProspect,NbCltC,
                    ResteAVisiter,Xvisit,
                    CltProspectCA,CltProspectQte,NewProspect,
                    RefCltVuCA,CCltVuCA,NouvCltCA,CltPerduCA,
                    RefCltVuQte,CCltVuQte,NouvCltQte,CltPerduQte,
                    TotalRef,TotalCourant,
                    TotalPerRefCA,realise,prevu,CAretenu,
                    TotalPerRefQte,QtePrevu,QteRealise,QteRet,extrapo]);
    end;
    Q.next;
  end;

  ferme(Q);
  if MiseANiveau=false then QTIns.free;

  { EVI / Chargement de l'arbre dans une TOB }
  Q2:=MOpenSql('SELECT QBR_NUMNOEUD, QBR_DATEDELAI,'+
                'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,'+
                'QBR_VALAXENIV5,QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,'+
                'QBR_VALAXENIV9,QBR_VALAXENIV10 '+
                'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"','BPInitArbre (InsertFeuillesMailles).',true);

  { EVI / Requête avec la date de la pièce pour QBPARBREDETAIL }
  if codeSqldetail <> '' then
  begin
    try
      QDet:=OpenSql(codeSqldetail,true);
    except
      SQLError := true;
      QDet:=nil;
      exit;
    end;
    while not QDet.eof do
    begin
      NumNoeudDetail:=NumNoeudDetail+1;

      QteC:=QDet.fields[0].asFloat;
      CA1:=QDet.fields[2].asFloat;
      CA2:=QDet.fields[3].asFloat;
      CA3:=QDet.fields[4].asFloat;
      CA4:=QDet.fields[5].asFloat;
      CA5:=QDet.fields[6].asFloat;
      CA6:=QDet.fields[7].asFloat;

      for i:=0 to niveau-1 do TabAxe[i]:=QDet.fields[8+i].asString;

      { EVI / Ajout de la date de la pièce - QBR_DATEPIECE }
      DatePiece:=QDet.fields[8+niveau+2].asDateTime;

      { EVI / iDate1900 pour DATEDELAI (génère 31/12/1899 avec 0) }
      DateDelai:=iDate1900;

      niveauX:=niveau;

      if OkDelai then
      begin
        niveauX:=niveau+1;
        DateDelai:=DateCDeb;
      end;

      while not Q2.eof do
      begin

        if DateDelai=Q2.fields[1].asDateTime then
        begin
          //Comparaison des axes
          for i:=0 to niveau-1 do
          begin
            if TabAxe[i]=Q2.fields[2+i].asString then AxeOK:=true
            else
            begin
              AxeOK:=false;
              break;
            end;
          end;
          if AxeOK=true then
          begin
            NumNoeudRef:=Q2.fields[0].asInteger;
            break
          end
        end;

        Q2.next;

      end;
      //Q2.next;
      Q2.First;

      QTDetIns.addValeur([codeSession,NumNoeudDetail,niveauX,
                      DateDelai,
                      QteC,0,QteC,0,CA1,0,CA1,0,CA2,0,CA2,0,CA3,0,CA3,0,CA4,0,CA4,0,
                      CA5,0,CA5,0,CA6,0,CA6,0,
                      datepiece,NumNoeudRef]);

      QDet.next;
    end;
  end;
  ferme(Q2);
  ferme(QDet);
  QTDetIns.free;
  Q2.free;
end;


//initialisation des feuilles de l'arbre
//dans le cas où on initialise une prevision (avec coefficient)
//si coeff1 = true alors on initialise toutes les feuilles avec coeff=1
//(dans le cas où on lance le calcul de prévision sans avoir lancer le calcul des coeffs avant)
//sinon on calule le coefficient
procedure InsertFeuilles(const codeSession,BPInitialise:hString;coeff1:boolean);
var Niveau,i,numNoeud,NumNoeudDetail,NumNoeudRef,NbValid:integer;
    SaiCmd,SaiCmdRef,CodeMaj:hString;
    DateCDeb,DateCFin,DateRefDeb,DateRefFin:TDateTime;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    ListMaille:TListMaille;
    Mi:TMaille;
    okDelai,okPremMaille,okDerMaille:Boolean;
begin
  SQLError := false;
  //niveau arbre
  Niveau:=ChercheNivMax(codeSession);

  numNoeud:=0;
  NumNoeudDetail:=0;

  ChercheSaisonCmd(codeSession,SaiCmd,SaiCmdRef);
  ChercheDateDDateFPeriode(codeSession,DateCDeb,DateCFin,DateRefDeb,DateRefFin);

  NbValid := ChercheNivMax(codesession);
  if DonneMethodeSession(codeSession)<>'2' then NbValid := NbValid+1;
  NbValid := ((NbValid*2)-1)*(NbValid-1);
  InitMoveProgressForm(nil, TraduireMemoire('Session : ')+codesession , TraduireMemoire('Veuillez patienter'), NbDeMaille(codesession,BPInitialise,DateRefDeb,DateRefFin) + NbValid, True, True);

  ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

  okDelai:=BPInitialise<>'0';

  if not OkSessionObjectif(codeSession)
  then CalculCoeff(codeSession);

  //si session definie par rapport à une saison de commande
  //on cherche les dates (date pivot) correspondant au début et à la fin de la saison
  { if SaiCmd<>''
  then DatesDebutFinSaisonCmd(SaiCmd,DateCDeb,DateCFin);
  if SaiCmdRef<>''
  then DatesDebutFinSaisonCmd(SaiCmdRef,DateRefDeb,DateRefFin);
  }
  ListMaille:=TListMaille.create();
  InitialiseListeMaille(VALEURI(BpInitialise),DateCDeb,DateCFin,
                       DateRefDeb,DateRefFin,ListMaille);

  okPremMaille:=false;
  okDerMaille:=false;
  for i:=0 to ListMaille.count-1 do
  begin
    if i=0
    then okPremMaille:=true;
    if i=ListMaille.count-1
    then okDerMaille:=true;

    if not MoveCurProgressForm(TraduireMemoire('Insertion des valeurs...')) then
    begin
      MoveCurProgressForm(TraduireMemoire('Annulation en cours...'));
      break;
    end;

    if SQLError = true then
    begin
      PGIError(TraduireMemoire('Erreur dans le traitement de la requête SQL.'), 'Session : '+codeSession);
      break;
    end;

    Mi:=TMaille(ListMaille[i]);
    NumNoeudRef:=numNoeud;
    InsertFeuillesMailles(okPremMaille,okDerMaille,
                         codeSession,BpInitialise,okDelai,coeff1,Niveau,SaiCmd,SaiCmdRef,
                         Mi.DateDebReference,Mi.DateFinReference,
                         Mi.DateDebCourante,Mi.DateFinCourante,
                         TabCodeAxe,TabNumValAxe,numNoeud,NumNoeudDetail,NumNoeudRef);

    okPremMaille:=false;
    okDerMaille:=false;
  end;
  freeListMaille(ListMaille);


  CodeMaj:='';
  //initialisation des coches
  if BPOkOrli
  then
  //-----------------> ORLI
  begin
    if BpInitialise<>'0'
    then codeMaj:=' ,QBS_INITDELAI="X",QBS_SESSIONECLAT="-",QBS_SESSIONECLATAI="-" '
    else codeMaj:=' ,QBS_INITDELAI="-",QBS_SESSIONECLAT="-",QBS_SESSIONECLATAI="-" ';
  end
  //ORLI <-----------------
  else
  begin
    if BpInitialise<>'0'
    then codeMaj:=' ,QBS_INITDELAI="X",QBS_BPINITIALISE="'+BPInitialise+
                '",QBS_SESSIONECLAT="-",QBS_SESSIONECLATAI="-" '
    else codeMaj:=' ,QBS_INITDELAI="-",QBS_BPINITIALISE="0",'+
                'QBS_SESSIONECLAT="-",QBS_SESSIONECLATAI="-" ';
  end;


  //mise à jour dans qbpsessionbp session init
  MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="X",QBS_OKMODIFSESSION="X" '+codeMaj+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPInitArbre (InsertFeuilles).');
  if not OkSessionObjectif(codeSession)
  then MExecuteSql('UPDATE QBPSESSIONBP SET QBS_INITCOEFF="X",QBS_INITPREVISION="-" '+
                   'WHERE QBS_CODESESSION="'+codeSession+'"',
                   'BPInitArbre (InsertFeuilles).');

end;


procedure InitArbre(const codeSession,BPInitialise:hString;coeff1,UpdateSession:boolean);
var NivMax,Noeud : integer;
numnoeudDetail : string;
Q :TQuery;
begin
  //Booléen pour mise à jour des sessions
  if UpdateSession = true then MiseANiveau:=true else MiseANiveau:=false;

  VideArbre(codeSession,MiseANiveau);

  //insertion des feuilles de l'arbre
  InsertFeuilles(codeSession,BPInitialise,coeff1);

  //construction de l'arbre à partir des feuilles
  MakeArbre(codeSession);

  if MiseANiveau=true then
  begin
    { EVI / Dans le cas d'une mise à niveau, récupération des noeuds ajoutés manuellement }
    NivMax := ChercheNivMaxSession(codeSession);
    Q :=OpenSQL('SELECT QBR_CODESESSION,QBR_NIVEAU,'+
                'QBR_DATEDELAI,'+
                'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                'QBR_DATEDELAI,QBR_NUMNOEUD'+
                ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NIVEAU="'+IntToStr(NivMax+1)+'"',True);

    While not Q.Eof do
    begin
      Noeud := Q.Fields[32].AsInteger;
      if not ExisteSQL('SELECT QBH_NUMNOEUDREF FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+
                       '" AND QBH_NUMNOEUDREF="'+IntToStr(Noeud)+'"')
      then
      begin
        numnoeudDetail:=IntToStr(BPIncrementeNumNoeudDetail(codeSession));
        MExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
                    'QBH_DATEDELAI,'+
                    'QBH_OP1,QBH_OPPRCT1,QBH_REF1,QBH_REFPRCT1,'+
                    'QBH_OP2,QBH_OPPRCT2,QBH_REF2,QBH_REFPRCT2,'+
                    'QBH_OP3,QBH_OPPRCT3,QBH_REF3,QBH_REFPRCT3,'+
                    'QBH_OP4,QBH_OPPRCT4,QBH_REF4,QBH_REFPRCT4,'+
                    'QBH_OP5,QBH_OPPRCT5,QBH_REF5,QBH_REFPRCT5,'+
                    'QBH_OP6,QBH_OPPRCT6,QBH_REF6,QBH_REFPRCT6,'+
                    'QBH_QTEC,QBH_QTECPRCT,QBH_QTEREF,QBH_QTEREFPRCT,'+
                    'QBH_DATEPIECE,QBH_NUMNOEUDREF) VALUES  ("'+ Q.Fields[0].AsString+'",'+numnoeudDetail+','+Q.Fields[1].AsString+
                    ',"'+USDATETIME(Q.Fields[2].AsDateTime)+
                    '",'+strFPoint(Q.Fields[3].AsFloat)+','+strFPoint(Q.Fields[4].AsFloat)+','+strFPoint(Q.Fields[5].AsFloat)+','+strFPoint(Q.Fields[6].AsFloat)+
                    ','+strFPoint(Q.Fields[7].AsFloat)+','+strFPoint(Q.Fields[8].AsFloat)+','+strFPoint(Q.Fields[9].AsFloat)+','+strFPoint(Q.Fields[10].AsFloat)+
                    ','+strFPoint(Q.Fields[11].AsFloat)+','+strFPoint(Q.Fields[12].AsFloat)+','+strFPoint(Q.Fields[13].AsFloat)+','+strFPoint(Q.Fields[14].AsFloat)+
                    ','+strFPoint(Q.Fields[15].AsFloat)+','+strFPoint(Q.Fields[16].AsFloat)+','+strFPoint(Q.Fields[17].AsFloat)+','+strFPoint(Q.Fields[18].asFloat)+
                    ','+strFPoint(Q.Fields[19].AsFloat)+','+strFPoint(Q.Fields[20].AsFloat)+','+strFPoint(Q.Fields[21].AsFloat)+','+strFPoint(Q.Fields[22].AsFloat)+
                    ','+strFPoint(Q.Fields[23].AsFloat)+','+strFPoint(Q.Fields[24].AsFloat)+','+strFPoint(Q.Fields[25].AsFloat)+','+strFPoint(Q.Fields[26].AsFloat)+
                    ','+strFPoint(Q.Fields[27].AsFloat)+','+strFPoint(Q.Fields[28].AsFloat)+','+strFPoint(Q.Fields[29].AsFloat)+','+strFPoint(Q.Fields[30].AsFloat)+
                    ',"'+USDATETIME(Q.Fields[31].AsDateTime)+'",'+IntToStr(Noeud)+')',
                    'BPInitArbre (InitArbre).');
      end;
      Q.next;
    end;
    ferme(Q);
  end;

  //Si le traitement a été annulé supression des données de l'arbre
  if not MoveCurProgressForm(TraduireMemoire('Validation...')) then
  begin
    ExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"');
    ExecuteSql('DELETE FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"');
    ExecuteSQL('UPDATE QBPSESSIONBP SET QBS_SESSIONINIT="-",QBS_BPINITIALISE="" WHERE QBS_CODESESSION="'+codeSession+'"');
    PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), 'Session : '+codeSession);
  end;

  FiniMoveProgressForm;
end;

end.
