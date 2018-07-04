unit BPVueClient;

interface

uses Sysutils,HCtrls,Hent1,
     UCtx,Uutil,
     BPFctSession,BPBasic;

procedure RemplitTableQBPCalculTmp(const codeSession:hString);

implementation

//remplit la table pour stat pivot
//dans le cas où l'histo est un objectif
procedure RemplitTableQBPCalculTmpParRapportObjectif(const codeSession,SessionObjectif,
          ChpSql,codeP,chpSqlLib,joinSqlLib,codeWhereClient,structure,
          codeWhereC,codeRestrictionP,ChpSqlLibArbre,joinSqlLibArbre,
          CodeWhereHNew,codeRestrictionPC,codeWhereH:hString;
          NivMax,NivMaxObj:integer;TabNumValAxe:array of hString);
var i:integer;
    codeE,codeSql,champClient:hString;
begin
 codeE:='';
 codeSql:='';

 //cherche champ client dans arbre
 champClient:=ChercheClientArbreSession(SessionObjectif,nivMax+1);

 for i:=1 to nivMax do
  begin
   codeSql:= codeSql+',QBR_VALAXENIV'+IntToStr(i);
  end;

 for i:=1 to nivMax do
  begin
   if TabNumValAxe[i]<>''
    then
     begin
      codeE:= codeE+' AND QBR_VALAXENIV'+IntToStr(i)+
              '=QBV_VALAXEP'+TabNumValAxe[i]+' ';
     end;
  end;

 //Histo
 //données objectif (qbparbre)
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            'SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+'" '+codeSql+
            ' '+ChpSqlLibArbre+',QBR_VALEURAXE,QBR_LIBVALAXE,'+
            'QBR_QTEC,QBR_OP1,1,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0'+
            ' FROM QBPARBRE '+joinSqlLibArbre+' WHERE QBR_CODESESSION="'+SessionObjectif+
            '" AND QBR_NIVEAU="'+IntToStr(NivMaxObj)+'"',
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 //Courant
 //donnes de pivot (qbppivot)
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 //Nouveaux clients
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            'SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,'+
            'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' AND NOT EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
            'WHERE C.QBR_CODESESSION="'+SessionObjectif+
            '" '+codeE+' AND '+champClient+'=QBV_VALAXEP13)'+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 //Clients Vus courant
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,0,0,0,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' AND QBV_VALAXEP13 in (SELECT distinct C.QBV_VALAXEP13 '+
            'FROM QBPPIVOT C WHERE C.QBV_CODESTRUCT="'+structure+
            '" '+CodeWhereHNew+codeRestrictionPC+') '+
            ' AND EXISTS (SELECT C.QBR_CODESESSION FROM QBPARBRE C '+
            'WHERE C.QBR_CODESESSION="'+SessionObjectif+
            '" '+codeE+' AND '+champClient+'=QBV_VALAXEP13)'+
           ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 //Clients Vus reference
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            'SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+'" '+codeSql+
            ' '+ChpSqlLibArbre+',QBR_VALEURAXE,QBR_LIBVALAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,QBR_QTEC,QBR_OP1,0,0,0 '+
            ' FROM QBPARBRE '+joinSqlLibArbre+' WHERE QBR_CODESESSION="'+SessionObjectif+
            '" AND QBR_NIVEAU="'+IntToStr(NivMaxObj)+'"'+
            ' AND EXISTS (SELECT QBV_QTE1 FROM QBPPIVOT '+
            'WHERE QBV_CODESTRUCT="'+structure+
            '" '+codeWhereH+codeE+' AND QBV_VALAXEP13'+
            '='+champClient+')',
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 //Prospect
 //données objectif (qbparbre)
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR) '+
            'SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+'" '+codeSql+
            ' '+ChpSqlLibArbre+',QBR_VALEURAXE,QBR_LIBVALAXE,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,QBR_QTEC,QBR_OP1,1 '+
            ' FROM QBPARBRE '+joinSqlLibArbre+' WHERE QBR_CODESESSION="'+SessionObjectif+
            '" AND QBR_NIVEAU="'+IntToStr(NivMaxObj)+
            '" AND EXISTS (SELECT X.QBA_CODEAXE FROM QBPAXE X WHERE '+
            'X.QBA_CODEAXE="CLIENT" AND X.QBA_VALEURAXE='+champClient+
            ' AND X.QBA_CATEGORIEC="P") ',
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');

 {
 //Clients perdus
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH) '+
            'SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+'" '+codeSql+
            ' '+ChpSqlLibArbre+',QBR_VALEURAXE,QBR_LIBVALAXE,'+
            'QBR_QTEC,QBR_OP1,1,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0'+
            ' FROM QBPARBRE '+joinSqlLibArbre+' WHERE QBR_CODESESSION="'+SessionObjectif+
            '" AND QBR_NIVEAU="'+IntToStr(NivMax+1)+'"'+
            ' AND EXISTS (SELECT QBP_BPSAICMD FROM QBPCLTPERDU WHERE '+
            'QBP_BPSAICMD=QBV_VALAXEP8 AND QBP_BPDIVCOM=QBV_VALAXEP9 '+
            'AND QBP_BPCLIENT=QBV_VALAXEP13 AND QBP_BPSOCIETE=QBV_VALAXEP12)',
            'RemplitTableQBPCalculTmpParRapportObjectif (BPVueClient).');
 }
end;


//remplit la table pour stat pivot
//dans le cas où l'histo est une période ou une saison de commande
procedure RemplitTableQBPCalculTmpParRapportPivot(const codeSession,
          ChpSql,codeP,chpSqlLib,joinSqlLib,codeWhereClient,structure,
          codeWhereH,codeWhereC,CodeWhereHNew,CodeWhereCNew,
          codeRestrictionP,codeRestrictionPC:hString);
begin
 //Histo
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1, '+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereH+codeRestrictionP+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');
 //Courant
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');

 //Nouveaux clients
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,'+
            'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,0 FROM QBPPIVOT P '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' AND NOT exists (SELECT C.QBV_VALAXEP13 '+
            'FROM QBPPIVOT C WHERE C.QBV_CODESTRUCT="'+structure+
            '" AND C.QBV_VALAXEP13=P.QBV_VALAXEP13 '+
            CodeWhereHNew+codeRestrictionPC+') '+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');

 //Clients Vus courant
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,0,0,0,0,0,0,0,0 FROM QBPPIVOT P '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            ' AND exists (SELECT C.QBV_VALAXEP13 '+
            'FROM QBPPIVOT C WHERE C.QBV_CODESTRUCT="'+structure+
            '" AND C.QBV_VALAXEP13=P.QBV_VALAXEP13 '+
            CodeWhereHNew+codeRestrictionPC+' ) '+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');

 //Clients Vus reference
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),0,0,0 FROM QBPPIVOT P '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereH+codeRestrictionP+
            ' AND  exists (SELECT C.QBV_CODESTRUCT'+
            ' FROM QBPPIVOT C WHERE C.QBV_CODESTRUCT="'+structure+ '"'+
            ' AND C.QBV_VALAXEP13=P.QBV_VALAXEP13 '+
            CodeWhereCNew+codeRestrictionPC+') '+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');

 //Clients perdus
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,0,0,0,'+
            '0,0,0,0,0,0,0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1,0,0,0,0,0 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereH+codeRestrictionP+
            ' AND EXISTS (SELECT QBP_BPSAICMD FROM QBPCLTPERDU WHERE '+
            'QBP_BPSAICMD=QBV_VALAXEP8 AND QBP_BPDIVCOM=QBV_VALAXEP9 '+
            'AND QBP_BPCLIENT=QBV_VALAXEP13 AND QBP_BPSOCIETE=QBV_VALAXEP12)'+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');

 //Clients Prospects
 MExecuteSql('INSERT INTO QBPCALCULTMP (QBU_CTX,QBU_CODESESSION'+ChpSql+
            ',QBU_CODECLIENT,QBU_LIBCODECLIENT,'+
            'QBU_QTEH,QBU_CAH,QBU_NBCLTH,'+
            'QBU_QTEC,QBU_CAC,QBU_NBCLTC,'+
            'QBU_QTEE,QBU_CAE,QBU_NBCLTE,'+
            'QBU_QTEN,QBU_CAN,QBU_NBCLTN,'+
            'QBU_QTEV,QBU_CAV,QBU_NBCLTV,'+
            'QBU_QTEP,QBU_CAP,QBU_NBCLTP,'+
            'QBU_QTEVH,QBU_CAVH,'+
            'QBU_QTECLTPR,QBU_CACLTPR,QBU_NBCLTPR)'+
            ' SELECT "'+ValeurContexte('qbpcalcultmp')+'","'+codeSession+
            '",'+codeP+chpSqlLib+',QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+
            '0,0,0,'+
            '0,0,0,'+
            '0,0,0,0,0,0,0,'+
            '0,0,0,0,0,0,0,SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+'+
            'QBV_QTET5+QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
            'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
            'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20),'+
            'SUM(QBV_QTE1),1 FROM QBPPIVOT '+joinSqlLib+
            ' LEFT JOIN QBPAXE AC on '+codeWhereClient+
            'AND AC.QBA_VALEURAXE=QBV_VALAXEP13 ' +
            ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeWhereC+codeRestrictionP+
            'AND EXISTS (SELECT X.QBA_CODEAXE FROM QBPAXE X WHERE '+
            'X.QBA_CODEAXE="CLIENT" AND X.QBA_VALEURAXE=QBV_VALAXEP13 '+
            ' AND X.QBA_CATEGORIEC="P") '+
            ' GROUP BY QBV_VALAXEP13,AC.QBA_LIBVALEURAXE,'+codeP+chpSqlLib,
            'RemplitTableQBPCalculTmpParRapportPivot (BPVueClient).');
end;


//remplit la table pour stat pivot
procedure RemplitTableQBPCalculTmp(const codeSession:hString);
var SessionObjectif:hString;
    i,NivMax,NivMaxObj,NivMaxLib:integer;
    ChpSql,codeP,chpSqlLib,joinSqlLib,ChpSqlLibArbre,joinSqlLibArbre:hString;
    CodeWhereC,CodeWhereH,CodeWhereHNew,CodeWhereCNew,CodeJoin:hString;
    codeRestrictionP,codeRestrictionPC,codeE,codeWhereClient:hString;
    structure,SaiCmd,SaiCmdRef:hString;
    TabCodeAxe:array [1..11] of hString;
    TabNumValAxe:array [1..11] of hString;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    OkCalculParRapportArbre:boolean;
begin
 if not OkSessionModifVue(codeSession)
  then exit;
 begintrans;
 MexecuteSql('DELETE FROM QBPCALCULTMP WHERE QBU_CTX="'+
             ValeurContexte('qbpcalcultmp')+
             '" AND QBU_CODESESSION="'+codeSession+'"',
            'RemplitTableQBPCalculTmp (BPVueClient).');

 //cherche le nombre de niveau definit
 nivMax:=ChercheNivMax(CodeSession);

 structure:=ChercheSessionStructure(codeSession);
 ChercheTabCodeAxeTabNumValAxe(CodeSession,TabCodeAxe,TabNumValAxe);

 ChercheSaisonCmd(CodeSession,SaiCmd,SaiCmdRef);
 ChercheDateDDateFPeriode(CodeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);

 if SaiCmd<>''
  then
   begin
    CodeWhereC:=' AND QBV_VALAXEP8="'+SaiCmd+'" ';
    CodeWhereCNew:=' AND C.QBV_VALAXEP8="'+SaiCmd+'" ';
   end
  else
   begin
    CodeWhereC:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateDebC)+
                   '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinC)+'" ';
    CodeWhereCNew:=' AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebC)+
                   '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinC)+'" ';
   end;
 if SaiCmdRef<>''
  then
   begin
    CodeWhereH:=' AND QBV_VALAXEP8="'+SaiCmdRef+'" ';
    CodeWhereHNew:=' AND C.QBV_VALAXEP8="'+SaiCmdRef+'" ';
   end
  else
   begin
    CodeWhereH:=' AND QBV_DATEPIVOT>="'+USDATETIME(DateDebRef)+
                '" AND QBV_DATEPIVOT<="'+USDATETIME(DateFinRef)+'" ';
    CodeWhereHNew:=' AND C.QBV_DATEPIVOT>="'+USDATETIME(DateDebRef)+
                   '" AND C.QBV_DATEPIVOT<="'+USDATETIME(DateFinRef)+'" ';
   end;
 
 codeRestrictionP:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,CodeJoin);
 codeRestrictionPC:=ChercheCodeRestrictionSession('C.','QBV_VALAXEP','QBV_',codeSession,CodeJoin);

 codeWhereClient:=' AC.QBA_CODEAXE="CLIENT" ';

 
 OkCalculParRapportArbre:=(ChercheSessionObjectif(codeSession,SessionObjectif));
 if OkCalculParRapportArbre
  then
   begin
    CodeWhereH:='';
    CodeWhereHNew:='';
   end;

 for i:=1 to NivMax do
  ChpSql:=ChpSql+',QBU_VALAXECU'+IntToStr(i);
 for i:=1 to NivMax do
  ChpSql:=ChpSql+',QBU_LIBVALAXECU'+IntToStr(i);
 codeE:='';
 for i:=1 to nivMax do
  begin
   if codeP=''
    then codeP:='QBV_VALAXEP'+TabNumValAxe[i+1]
    else codeP:=codeP+',QBV_VALAXEP'+TabNumValAxe[i+1];
   codeE:=codeE+' AND QBV_VALAXEP'+TabNumValAxe[i+1]+
          '=C.QBV_VALAXEP'+TabNumValAxe[i+1]+' ';
  end;

 NivMaxLib:=NivMax+1;

 //code pour avoir les libelles
 ChercheCodeSqlChpAxeLibAvecPivot(NivMax,'',TabCodeAxe,TabNumValAxe,
                                  chpSqlLib,joinSqlLib);
 ChercheCodeSqlChpAxeLib(NivMaxLib,TabCodeAxe,ChpSqlLibArbre,joinSqlLibArbre);

 NivMaxObj:=NivMax+1;
 if OkSessionObjectif(codeSession) and (CodeWhereH='')
  then
   begin
    SessionObjectif:=codeSession;
    OkCalculParRapportArbre:=true;
    NivMaxObj:=NivMax;
   end;

 if OkCalculParRapportArbre
  then RemplitTableQBPCalculTmpParRapportObjectif(codeSession,SessionObjectif,
                    ChpSql,codeP,chpSqlLib,joinSqlLib,codeWhereClient,structure,
                    codeWhereC,codeRestrictionP,ChpSqlLibArbre,joinSqlLibArbre,
                    CodeWhereHNew,codeRestrictionPC,codeWhereH,
                    NivMax,NivMaxObj,TabNumValAxe)
  else RemplitTableQBPCalculTmpParRapportPivot(codeSession,
                    ChpSql,codeP,chpSqlLib,joinSqlLib,codeWhereClient,structure,
                    codeWhereH,codeWhereC,CodeWhereHNew,CodeWhereCNew,
                    codeRestrictionP,codeRestrictionPC);
 
 ModifSessionVue(CodeSession,'-','X');
 committrans;
end;

end.
