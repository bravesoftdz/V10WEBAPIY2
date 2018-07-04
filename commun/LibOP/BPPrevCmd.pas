unit BPPrevCmd;

interface

uses HCtrls,classes,Sysutils,Hent1;


procedure CalculPrevCmd(const codeSession,contexte:hString);

implementation

uses BPMaille,UDatamem,CstCommun,UCtx,LesDatamems,Udate,
     USema,UUtil,BPFctSession,UBasic,BPUtil;

//remplit le dm avec le prévu - le commandé
procedure DmPrevuCmd(niv:integer;Mi:TMaille);
var codeSql:hString;
begin
 //cas rothelec qte prévu - qté cmd
  codeSql:='SELECT DISTINCT P.QBV_VALAXEP1,P.QBV_VALAXEP2,"'+IntToStr(Mi.code)+
          '",P.QBV_VALAXEP18,".",".",".",'+
          '(SELECT SUM(QBB_QTEBPT1) FROM QBPBUDGETPREV '+
          'WHERE QBB_VALAXEBP2=P.QBV_VALAXEP2 AND QBB_DATEBP>="'+
          USDATETIME(Mi.DateDebCourante)+'" '+
          'AND QBB_DATEBP<"'+USDATETIME(Mi.DateFinCourante)+
          '") AS PREVU,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'+
          '(SELECT SUM(P1.QBV_QTET1) '+
          'FROM QBPPIVOT P1 WHERE P1.QBV_VALAXEP2=P.QBV_VALAXEP2 and '+
          'P1.qbv_datepivot>="'+USDATETIME(Mi.DateDebCourante)+
          '" and P1.qbv_datepivot<"'+USDATETIME(Mi.DateFinCourante)+'") AS CMD,'+
          '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 '+
          ' FROM QBPPIVOT P ';

 dm_ajouteSql(idm_BPTmpPrevCmd,codeSql);
end;

procedure InsertDmQtranche(codeSession,contexte:hString;ListMaille:TListMaille);
var enr:TEnreg;
    Pseq:Tdmpseq;
    maille:integer;
    tranche,saison,article,codearticle,coloris,fs,magasin:hString;
    numtrandetd:integer;
    dateT:TDateTime;
    qteT1,qteT2,qteT3,qteT4,qteT5,qteT6,qteT7,qteT8,qteT9,qteT10:double;
    qteT11,qteT12,qteT13,qteT14,qteT15,qteT16,qteT17,qteT18,qteT19,qteT20:double;
    qteTC1,qteTC2,qteTC3,qteTC4,qteTC5,qteTC6,qteTC7,qteTC8,qteTC9,qteTc10:double;
    qteTC11,qteTC12,qteTC13,qteTC14,qteTC15,qteTC16,qteTC17,qteTC18,qteTC19,qteTC20:double;
    qteTot:double;
    codeSql,origine:hString;
    Y,M,D:word;
begin
 tranche:='';
 numtrandetd:=1;

 //parcours datamem
 if dm_pseqAll(idm_BPTmpPrevCmd, Pseq) = 0 then
  while dm_pseqlect(Pseq, enr) = 0 do
   begin
    maille:=Valeuri(enr.ch(BPTMPPREVCMD_MAILLE));
    saison:=enr.ch(BPTMPPREVCMD_SAISON);
    article:=enr.ch(BPTMPPREVCMD_ARTICLE);
    codearticle:=copy(article,1,18);
    qteT1:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET1)));
    qteT2:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET2)));
    qteT3:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET3)));
    qteT4:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET4)));
    qteT5:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET5)));
    qteT6:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET6)));
    qteT7:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET7)));
    qteT8:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET8)));
    qteT9:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET9)));
    qteT10:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET10)));
    qteT11:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET11)));
    qteT12:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET12)));
    qteT13:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET13)));
    qteT14:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET14)));
    qteT15:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET15)));
    qteT16:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET16)));
    qteT17:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET17)));
    qteT18:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET18)));
    qteT19:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET19)));
    qteT20:=round(Valeur(enr.ch(BPTMPPREVCMD_QTET20)));    
    qteTC1:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC1)));
    qteTC2:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC2)));
    qteTC3:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC3)));
    qteTC4:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC4)));
    qteTC5:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC5)));
    qteTC6:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC6)));
    qteTC7:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC7)));
    qteTC8:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC8)));
    qteTC9:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC9)));
    qteTC10:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC10)));
    qteTC11:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC11)));
    qteTC12:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC12)));
    qteTC13:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC13)));
    qteTC14:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC14)));
    qteTC15:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC15)));
    qteTC16:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC16)));
    qteTC17:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC17)));
    qteTC18:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC18)));
    qteTC19:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC19)));
    qteTC20:=round(Valeur(enr.ch(BPTMPPREVCMD_QTETC20)));
    DateT:=MailleDonneDate(maille,ListMaille);
    numtrandetd:=numtrandetd+1;
    decodeDate(DateT,Y,M,D);
    if ContextBP in [0,2,3]
     then
      begin
       tranche:=codeSession;
       origine:=codeSession;
      end
     else tranche:='P'+IntToStr(Y)+MetZero(IntToStr(M),2)+MetZero(IntToStr(D),2);
    coloris:=enr.ch(BPTMPPREVCMD_COLORIS);
    fs:=enr.ch(BPTMPPREVCMD_FS);
    magasin:=enr.ch(BPTMPPREVCMD_MAGASIN);
    qteTot:=(qteT1-qteTC1)+(qteT2-qteTC2)+(qteT3-qteTC3)+(qteT4-qteTC4)+
            (qteT5-qteTC5)+(qteT6-qteTC6)+(qteT7-qteTC7)+(qteT8-qteTC8)+
            (qteT9-qteTC9)+(qteT10-qteTC10)+(qteT11-qteTC11)+(qteT12-qteTC12)+
            (qteT13-qteTC13)+(qteT14-qteTC14)+(qteT15-qteTC15)+(qteT16-qteTC16)+
            (qteT17-qteTC17)+(qteT18-qteTC18)+(qteT19-qteTC19)+(qteT20-qteTC20);
    if (qteTot>0)
     then
      begin
       if ContextBP in [0,2,3]
        then MExecuteSql('INSERT INTO WCBNBESOIN (WBE_CTX,WBE_ARTICLE,WBE_CODEARTICLE,'+
                         'WBE_DEPOT,WBE_QUANTITE,WBE_DATELIVRAISON,WBE_TRANCHE,'+
                         'WBE_ORIGINE) VALUES '+
                         '("'+contexte+'","'+article+'","'+codearticle+
                         '","'+magasin+'","'+STRFPOINT(qteTot)+'","'+
                         USDATETIME(DateT)+'","'+tranche+'","'+origine+'")',
                         'BPrevCmd(InsertDmWCBNBESOIN)')
        else
         begin
          //insertion dans qtranche
          MExecuteSql('INSERT INTO QTRANCHE (QT_CTX,QT_SAISON,QT_ARTTECH,QT_TRANCHE,'+
            'QT_QTETRANCHE,QT_QTEPREVUE,QT_DATEFIN,QT_TYPEBESOIN) VALUES '+
            '("'+ValeurContexte('qtranche')+'","'+saison+'","'+article+
            '","'+(tranche)+'","'+STRFPOINT(qteTot)+'","'+
            STRFPOINT(qteTot)+'","'+USDATETIME(DateT)+'","2")',
            'BPrevCmd(InsertDmQtranche)');
          //insertion dans qtrandetd
          MExecuteSql('INSERT INTO QTRANDETD (QTD_CTX,QTD_SAISON,QTD_ARTTECH,QTD_TRANCHE,'+
            'QTD_NUMDETTR,QTD_COLORIS,QTD_FS,QTD_MAGASIN,'+
            'QTD_QTC1,QTD_QTC2,QTD_QTC3,QTD_QTC4,QTD_QTC5,QTD_QTC6,QTD_QTC7,'+
            'QTD_QTC8,QTD_QTC9,QTD_QTC10,QTD_QTC11,QTD_QTC12,QTD_QTC13,'+
            'QTD_QTC14,QTD_QTC15,QTD_QTC16,QTD_QTC17,QTD_QTC18,QTD_QTC19,QTD_QTC20,'+
            'QTD_QTP1,QTD_QTP2,QTD_QTP3,QTD_QTP4,QTD_QTP5,QTD_QTP6,QTD_QTP7,'+
            'QTD_QTP8,QTD_QTP9,QTD_QTP10,QTD_QTP11,QTD_QTP12,QTD_QTP13,'+
            'QTD_QTP14,QTD_QTP15,QTD_QTP16,QTD_QTP17,QTD_QTP18,'+
            'QTD_QTP19,QTD_QTP20) VALUES '+
            '("'+ValeurContexte('qtranche')+'","'+saison+'","'+article+
            '","'+(tranche)+'","'+STRFPOINT(numtrandetd)+
            '","'+coloris+'","'+fs+'","'+magasin+'","'+STRFPOINT(qteT1)+'","'+STRFPOINT(qteT2)+
            '","'+STRFPOINT(qteT3)+'","'+STRFPOINT(qteT4)+'","'+STRFPOINT(qteT5)+
            '","'+STRFPOINT(qteT6)+'","'+STRFPOINT(qteT7)+
            '","'+STRFPOINT(qteT8)+'","'+STRFPOINT(qteT9)+'","'+STRFPOINT(qteT10)+
            '","'+STRFPOINT(qteT11)+'","'+STRFPOINT(qteT12)+
            '","'+STRFPOINT(qteT13)+'","'+STRFPOINT(qteT14)+'","'+STRFPOINT(qteT15)+
            '","'+STRFPOINT(qteT16)+'","'+STRFPOINT(qteT17)+
            '","'+STRFPOINT(qteT18)+'","'+STRFPOINT(qteT19)+'","'+STRFPOINT(qteT20)+
            '","'+STRFPOINT(qteT1)+'","'+STRFPOINT(qteT2)+
            '","'+STRFPOINT(qteT3)+'","'+STRFPOINT(qteT4)+'","'+STRFPOINT(qteT5)+
            '","'+STRFPOINT(qteT6)+'","'+STRFPOINT(qteT7)+
            '","'+STRFPOINT(qteT8)+'","'+STRFPOINT(qteT9)+'","'+STRFPOINT(qteT10)+
            '","'+STRFPOINT(qteT11)+'","'+STRFPOINT(qteT12)+
            '","'+STRFPOINT(qteT13)+'","'+STRFPOINT(qteT14)+'","'+STRFPOINT(qteT15)+
            '","'+STRFPOINT(qteT16)+'","'+STRFPOINT(qteT17)+
            '","'+STRFPOINT(qteT18)+'","'+STRFPOINT(qteT19)+'","'+STRFPOINT(qteT20)+'")',
            'BPrevCmd(InsertDmQtranche)');
         end;

       end;
   end;
end;


procedure CalculPrevActualisee();
var Pseq:TdmPseq;
    L:Tenreg;
    saison,article,saisonSvg,articleSvg:hString;
  //  maille:integer;
    qte,qteX:double;
begin
{ ok := dm_pseq(idm_BPTmpPrevCmd, [saison,article], pseq);
 while (ok=0) and (dm_pseqlect(Pseq, L) = 0) do
}
 saisonSvg:='';
 articleSvg:='';
 //valeur de surplus à affecter
 qteX:=0;
 if dm_pseqAll(idm_BPTmpPrevCmd, Pseq) = 0
  then
   while (dm_pseqlect(Pseq, L) = 0) do
    begin
     saison:=L.ch(BPTMPPREVCMD_SAISON);
     article:=L.ch(BPTMPPREVCMD_ARTICLE);
   //  maille:=ValeurI(L.ch(BPTMPPREVCMD_MAILLE));
     qte:=Valeur(L.ch(BPTMPPREVCMD_QTET1));
     if qteX=0
      then
       begin
        if qte<0
         then qteX:=-qte;
       end
      else
       begin

       end;
     if (saison<>saisonSvg) or (article<>articleSvg)
      then
       begin
        saisonSvg:=saison;
        articleSvg:=article;

       end;


     // dm_modifchcour(idm_BPTmpPrevCmd,BPTMPPREVCMD_QTET1,'0');


  end;
end;


procedure InsertRAL;
begin
 //insertion dans qtranche
 MExecuteSql('INSERT INTO QTRANCHE (QT_CTX,QT_SAISON,QT_ARTTECH,QT_TRANCHE,'+
            'QT_QTETRANCHE,QT_QTEPREVUE,QT_DATEFIN,QT_TYPEBESOIN) '+
            ' SELECT '+ValeurContexte('qtranche')+',QBV_VALAXEP1,QBV_VALAXEP2,'+
            'QBV_VALAXEP18,'+
            'SUM(QBV_QTET1)-SUM(QBV_QTE5),SUM(QBV_QTET1)-SUM(QBV_QTE5),'+
            'QBV_DATEPIVOT,"1" FROM QBPPIVOT where QBV_VALAXEP18 like "_%" '+
            'GROUP BY QBV_VALAXEP1,QBV_VALAXEP2,QBV_VALAXEP18,QBV_DATEPIVOT '+
            'HAVING SUM(QBV_QTET1)-SUM(QBV_QTE5)>0','BPrevCmd(InsertRAL)');

 //insertion dans qtrandetd
 MExecuteSql('INSERT INTO QTRANDETD (QTD_CTX,QTD_SAISON,QTD_ARTTECH,QTD_TRANCHE,'+
            'QTD_NUMDETTR,QTD_COLORIS,QTD_FS,QTD_MAGASIN,'+
            'QTD_QTC1,QTD_QTC2,QTD_QTC3,QTD_QTC4,QTD_QTC5,QTD_QTC6,QTD_QTC7,'+
            'QTD_QTC8,QTD_QTC9,QTD_QTC10,QTD_QTC11,QTD_QTC12,QTD_QTC13,'+
            'QTD_QTC14,QTD_QTC15,QTD_QTC16,QTD_QTC17,QTD_QTC18,QTD_QTC19,QTD_QTC20,'+
            'QTD_QTP1,QTD_QTP2,QTD_QTP3,QTD_QTP4,QTD_QTP5,QTD_QTP6,QTD_QTP7,'+
            'QTD_QTP8,QTD_QTP9,QTD_QTP10,QTD_QTP11,QTD_QTP12,QTD_QTP13,'+
            'QTD_QTP14,QTD_QTP15,QTD_QTP16,QTD_QTP17,QTD_QTP18,'+
            'QTD_QTP19,QTD_QTP20)  '+
            ' SELECT '+ValeurContexte('qtrandetd')+',QBV_VALAXEP1,QBV_VALAXEP2,'+
            'QBV_VALAXEP18,'+
            'TRIM(STR(MONTH(QBV_DATEPIVOT)))||"-"||TRIM(STR(DAY(QBV_DATEPIVOT))),'+
            '".",".",".",SUM(QBV_QTET1)-SUM(QBV_QTE5),0,'+
            '0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,SUM(QBV_QTET1)-SUM(QBV_QTE5)'+
            ',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 '+
            'FROM QBPPIVOT where QBV_VALAXEP18 like "_%"'+
            'GROUP BY QBV_VALAXEP1,QBV_VALAXEP2,QBV_VALAXEP18,QBV_DATEPIVOT '+
            'HAVING SUM(QBV_QTET1)-SUM(QBV_QTE5)>0','BPrevCmd(InsertRAL)');

  //faire attention au numdet
end;

//contexte <> '' avec PGIENT
//sinon contexte=''
procedure CalculPrevCmd(const codeSession,contexte:hString);
var ListMaille:TListMaille;
    i,niv:integer;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    types:hString;
    Mi:TMaille;
    DmBP: TDataMem;
begin
 //
 MExecuteSql('DELETE FROM QBPBUDGETPREV WHERE QBB_CODESESSION="'+codeSession+'"',
             'BPrevCmd(CalculPrevCmd)');

 begintrans;
 //suppression dans qtranche et qtrandetd des enr venant de la prévision
 MExecuteSql('DELETE FROM QTRANCHE WHERE '+WhereCtx(['qtranche']),
             'BPrevCmd(CalculPrevCmd)');
 MExecuteSql('DELETE FROM QTRANDETD WHERE '+WhereCtx(['qtrandetd']),
             'BPrevCmd(CalculPrevCmd)');


 //cherche les dates de la session
 ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);
 //initialise la liste des mailles pour la session
 ListMaille:=TListMaille.create();
 SessionCalculParDelai(codeSession,types);
 types:='4';
 InitialiseListeMaille(VALEURI(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);
 niv:=ChercheNivMax(codeSession);
 dm_init(idm_BPTmpPrevCmd,0,['A','A','A'],true,
         [rc_SAISON,rc_MATPROD,rc_MAILLE],DmBP);

 //parcours des mailles
 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);
   //remplit le dm avec le prévu - le commandé
   //on ne prend que les périodes < à la date du stock
   if Mi.DateFinCourante>dh_date(DateHeureStock)
    then DmPrevuCmd(niv,Mi);
  end;

 //calcul la prevision nette
 //càd affectation du surplus suivant la methode choisie
 CalculPrevActualisee();
 //insertion dans la table qtranche
 InsertDmQtranche(codeSession,contexte,ListMaille);

 //reste à livré
 if not(ContextBP in [0,2,3])
  then InsertRAL;

 freeListMaille(ListMaille);
 LibereUnDm(idm_BPTmpPrevCmd);

 CommitTrans;

 libereLesDm([idm_tranche,idm_trandetd]);

 AvertirUser('BP');
 InValideCalculBesSattribTranche; // Ajout Daniel 10/02/05

 //MAJ du paramètre derniere validation vers SCM
 if not (ContextBP in [0,2,3])
  then ModifParam(ps_BPDateValidationSCM,DateTimeToSTr(now)+' (session : '+codeSession+')');
end;

end.
