unit BPCubeTmp;

interface

uses HEnt1;

procedure RemplitTableQBPCubeTmpOrli(const codeSession:hString;
          DateDebCourante,DateFinCourante:TDateTime);
{ EVI / Ajout Date Edition Debut, Date Edition Fin }
procedure RemplitTableQBPCubeTmpPgi(const codeSession,BaseSQL,UserLimit:hString;
          DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin:TDateTime);
//CEGID-CCMX le 07/11/2006
procedure RemplitTableQBPCubeTmpPgiTYPEDA(
          codeSession,codeChpSql,champs,ChpTotal,joinChpSql,
          codeRestriction,UserLimit,SQLUserLimit:hstring;
          TabCodeAxe:array of hstring;
          DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin:TDateTime);
implementation

uses SysUtils,Classes,HMsgBox,StrUtils,ed_tools,HCtrls,utilPGI,
     {$IFNDEF EAGLCLIENT}    //CEGID-CCMX le 07/11/2006 DEBUT
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else}utob,{$ENDIF}    //CEGID-CCMX le 07/11/2006 FIN
     {$IFNDEF MODE}{$IFDEF GCGC}paramsoc,{$ENDIF}{$ENDIF}
     UUtil,UCtx,BPUtil,BPFctSession,BPBasic,BPMaille;

//-----------------> ORLI
procedure RemplitTableQBPCubeTmpOrli0(const codeSession,codeChpSql,codeP,
          codeChpSqlLib,
          structure,codeRestrictionP,chpSqlLib,joinSqlLib:hString;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
          TabCodeAxe,TabNumValAxe:array of hString;
          SaiCmd,SaiRef:hString);
var _SessionObj,codeWhereSaiRef,codeWhereSaiC:hString;
    codesqlArbre:hString;
    nivMax,i:integer;
    ListMaille:TListMaille;
    Mi:TMaille;
    types,chpArbreLib,joinArbreLib,codefiltre,codeChpVal,codeDate:hString;
    OKSessionOrgObjectif: boolean;
begin
 ListMaille:=TListMaille.create();
 SessionCalculParDelai(codeSession,types);
 if SessionEclateeParDelai(codeSession)
  then types:='1';

 InitialiseListeMaille(VALEURI(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

 codeWhereSaiC:='';
 codeWhereSaiRef:='';
 if SaiCmd<>''
  then codeWhereSaiC:=' AND QBV_VALAXEP8="'+SaiCmd+'" ';
 if SaiRef<>''
  then codeWhereSaiRef:=' AND QBV_VALAXEP8="'+SaiRef+'" ';

 OKSessionOrgObjectif := ChercheSessionObjectif(codesession,_SessionObj);
 //pour chaque maille
 for i:=0 to ListMaille.count-1 do
  begin
   Mi:=TMaille(ListMaille[i]);
   if ListMaille.count=1 then codeDate:=''
                         else
   begin          // au moins 2 mailles
     codeDate:=' AND QBV_DATEPIVOT>="'+ USDATETIME(Mi.DateDebReference)+
                '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinReference)+'" ';
     if (i=0)  and (SaiCmd<>'') then //premiere maille
       codeDate:=' AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinReference)+'" ';
     if (i=ListMaille.count-1) and (SaiCmd<>'') then  //derniere maille
      codeDate:=' AND QBV_DATEPIVOT>="'+USDATETIME(Mi.DateDebReference)+'" ';
   end;

   //*************** insertion réalisé
   //on prend dans la table qbppivot
   //pour la structure et les dates courantes
   MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSqlLib+
              ',QBQ_DATECT,'+
              'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
              'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+chpSqlLib+',"'+
              USDATETIME(Mi.DateDebCourante)+'",'+
              'SUM(QBV_QTE1),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" AND QBV_DATEPIVOT>="'+
              USDATETIME(Mi.DateDebCourante)+
              '" AND QBV_DATEPIVOT<="'+USDATETIME(Mi.DateFinCourante)+
              '" '+codeWhereSaiC+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCubeTmp (RemplitTableQBPCubeTmpOrli0).');
   //*************** insertion histo

   //si la session est une prevision
   //et qu'elle n'est pas initialisée par rapport à un objectif
   if (not OKSessionOrgObjectif)
    then MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSqlLib+
              ',QBQ_DATECT,'+
              'QBQ_CAHISTO,QBQ_CAREALISE,QBQ_CAPREVU,'+
              'QBQ_HISTO,QBQ_REALISE,QBQ_PREVU) '+
              ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
              '",'+codeP+chpSqlLib+
              ',"'+USDATETIME(Mi.DateDebCourante)+'",SUM(QBV_QTE1),0,0,'+
              'SUM(QBV_QTET1+QBV_QTET2+QBV_QTET3+QBV_QTET4+QBV_QTET5+'+
              'QBV_QTET6+QBV_QTET7+QBV_QTET8+QBV_QTET9+QBV_QTET10+'+
              'QBV_QTET11+QBV_QTET12+QBV_QTET13+QBV_QTET14+QBV_QTET15+'+
              'QBV_QTET16+QBV_QTET17+QBV_QTET18+QBV_QTET19+QBV_QTET20)'+
              ',0,0 FROM QBPPIVOT '+joinSqlLib+
              ' WHERE QBV_CODESTRUCT="'+structure+'" '+codeDate+codeWhereSaiRef+codeRestrictionP+
              ' GROUP BY '+codeP+',QBV_DATEPIVOT'+chpSqlLib,
              'BPCubeTmp (RemplitTableQBPCubeTmpOrli0).');
  end;
  freeListMaille(ListMaille);
  //*************** insertion prévu
  nivMax:=ChercheNivMax(codeSession);
  if SessionInitPrev(codeSession) then
  begin
    nivMax:=ChercheNivMaxSession(codeSession);
    if SessionDelai(codeSession) then nivMax:=nivMax-1;
  end;
  if SessionInitCoeff(codeSession) then codeChpVal:='QBR_QTERETENUE'
                                   else codeChpVal:='QBR_QTEC';
  if (SessionDelai(codeSession)) and (OkSessionObjectif(codeSession))
           then nivMax:=nivMax+1;
  codesqlArbre:='';
  for i:=1 to NivMax-1 do
  begin
    if tabCodeAxe[i]='' then continue;
    if codesqlArbre=''  then codesqlArbre:=' QBR_VALAXENIV'+IntToStr(i)
                        else codesqlArbre:=codesqlArbre+',QBR_VALAXENIV'+IntToStr(i);
  end;
  if tabCodeAxe[NivMax]<>'' then // le dernier niveau est enlevé de la requete : cas de l'axe COLORIS qui est 2 fois
  begin
    if not SessionDelai(codeSession)
           then codesqlArbre:=codesqlArbre+',QBR_VALEURAXE';
    if SessionDelai(codeSession) and SessionInitPrev(codeSession)
      then
    begin
      codesqlArbre:=codesqlArbre+',QBR_VALAXENIV'+IntToStr(nivmax);//',QBR_VALEURAXE';
      nivMax:=nivMax+1;
    end;
  end;
  ChercheCodeSqlChpAxeLib(nivMax,TabCodeAxe,chpArbreLib,joinArbreLib);

  MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+codeChpSqlLib+',QBQ_DATECT,'+
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
            '0,'+codeChpVal+',0,'+
            '0,QBR_OP2,0,'+
            '0,QBR_OP3,0,'+
            '0,QBR_OP4,0,'+
            '0,QBR_OP5,0,'+
            '0,QBR_OP6,0 '+
            ' FROM QBPARBRE '+joinArbreLib+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NIVEAU="'+IntToStr(NivMax)+'" '+codefiltre,
            'BPCubeTmp (RemplitTableQBPCubeTmpOrli0).');
end;
//ORLI <-----------------


//remplit table qbpcubetmp
//dans le cas orli
//-----------------> ORLI
procedure RemplitTableQBPCubeTmpOrli(const codeSession:hString;
          DateDebCourante,DateFinCourante:TDateTime);
var codeRestrictionP:hString;
    codeChpSql,codeP,structure:hString;
    codeChpSqlLib,chpSqlLib,joinSqlLib:hString;
    i:integer;
    nivMax:integer;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    SaiCmd,SaiRef,codeJoinR:hString;
begin
  if (not OkSessionModif(codeSession)) and
    (ExisteSql('SELECT QBQ_CODESESSION FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+codeSession+'"'))
       then exit;

  //vide la table temporaire
  MExecuteSql('DELETE FROM QBPCUBETMP WHERE '+WhereCtx(['qbpcubetmp'])+
            ' AND QBQ_CODESESSION="'+codeSession+'"',
            'BPCubeTmp (RemplitTableQBPCubeTmpOrli).');
  structure:=ChercheSessionStructure(codeSession);
  ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

  //cherche les codes pour le sql
  codeP:='';
  codeChpSql:='';

 //cube avec axe session
  nivMax:=ChercheNivMax(codeSession);
  if SessionInitPrev(codeSession) then
  begin
    //niv article
    if not(existCodeAxe(TabCodeAxe,'ARTICLE')) then
    begin
      TabCodeAxe[nivMax+2]:='ARTICLE';
      TabNumValAxe[nivMax+2]:='2';
      NivMax:=nivMax+1;
    end;
    //niv coloris
    if GestionBPColoris and not(existCodeAxe(TabCodeAxe,'COLORIS')) then
    begin
      TabCodeAxe[nivMax+2]:='COLORIS';
      TabNumValAxe[nivMax+2]:='3';
      nivMax:=nivMax+1;
    end;
    //niv FS
    if GestionBPFS and not(existCodeAxe(TabCodeAxe,'FS')) then
    begin
      TabCodeAxe[nivMax+2]:='FS';
      TabNumValAxe[nivMax+2]:='4';
      nivMax:=nivMax+1;
    end;
    //niv Magasin
    if GestionBPMagasin and not(existCodeAxe(TabCodeAxe,'MAGASIN')) then
    begin
      TabCodeAxe[nivMax+2]:='MAGASIN';
      TabNumValAxe[nivMax+2]:='7';
    end;
    nivMax:=ChercheNivMaxSession(codeSession);
    if SessionDelai(codeSession) then nivMax:=nivMax-1;
  end;
  for i:=1 to nivMax do
  begin
    if TabNumValAxe[i+1]='' then continue;
    if codeP='' then
    begin
      codeP:='QBV_VALAXEP'+TabNumValAxe[i+1];
      codeChpSql:='QBQ_VALAXECT'+TabNumValAxe[i+1];
    end
      else
    begin
      codeP:=codeP+',QBV_VALAXEP'+TabNumValAxe[i+1];
      codeChpSql:=codeChpSql+',QBQ_VALAXECT'+TabNumValAxe[i+1];
    end;
  end;
  //libellé
  codeChpSqlLib:=codeChpSql;
  for i:=1 to nivMax do
    if (TabNumValAxe[i+1]<>'') then codeChpSqlLib:=codeChpSqlLib+',QBQ_LIBVALAXECT'+TabNumValAxe[i+1];
  codeRestrictionP:=ChercheCodeRestrictionSession('','QBV_VALAXEP','QBV_',codeSession,codeJoinR);

  //cherche dates periode courante et reference de la session
  ChercheDateDDateFPeriode(codeSession,
                           DateDebC,DateFinC,DateDebRef,DateFinRef);
  //cherche saison de commande et de reference
  ChercheSaisonCmd(codeSession,SaiCmd,SaiRef);

 { //si session definie par rapport à une saison de commande
 //on cherche les dates (date pivot) correspondant au début et à la fin de la saison
 if SaiCmd<>''
  then DatesDebutFinSaisonCmd(SaiCmd,DateDebC,DateFinC);
 if SaiRef<>''
  then DatesDebutFinSaisonCmd(SaiRef,DateDebRef,DateFinRef);
}

 //code pour avoir les libelles
  ChercheCodeSqlChpAxeLibAvecPivot(NivMax,'',TabCodeAxe,TabNumValAxe,
                                  chpSqlLib,joinSqlLib);

  //remplit table tmp cube
  //avec l'histo, le realisé et le prévu
  RemplitTableQBPCubeTmpOrli0(codeSession,codeChpSql,codeP,
           codeChpSqlLib,
           structure,codeRestrictionP,chpSqlLib,joinSqlLib,
           DateDebC,DateFinC,DateDebRef,DateFinRef,TabCodeAxe,TabNumValAxe,
           SaiCmd,SaiRef);
  ModifSession(codeSession,'-');
end;
//ORLI <-----------------

//*****************************************************************************
//
// remplit cube avec PGI
//CEGID-CCMX : ajout sTYPEDA
procedure RemplitTableQBPCubeTmpPrevuObjDelaiMaille(const codeSession,BaseSQL,codeChpSql:hString;
          UserLimit:hString;Mi:TMaille;
          TabCodeAxe:array of hString; sTypeDA : string = '');
var nivMax,i:integer;
    codesqlArbre,chpArbreLib,joinArbreLib,codeWhere,SQLUserLimit,Values:hString;
begin
  nivMax:=ChercheNivMax(codeSession);
  codeWhere:='';
  if SessionDelai(codeSession) then
  begin
    nivMax:=nivMax+1;
    codeWhere:=' AND QBR_DATEDELAI>="'+USDATETIME(Mi.DateDebCourante)+
               '" AND QBR_DATEDELAI<="'+USDATETIME(Mi.DateFinCourante)+'"';
  end;

  codesqlArbre:='';

  for i:=1 to NivMax-1 do
  begin
    if codesqlArbre='' then codesqlArbre:=' QBR_VALAXENIV'+IntToStr(i)
    else codesqlArbre:=codesqlArbre+',QBR_VALAXENIV'+IntToStr(i);
  end;

  if (not SessionDelai(codeSession)) then
    if (codesqlArbre='') then codesqlArbre:=' QBR_VALEURAXE '
      else codesqlArbre:=codesqlArbre+',QBR_VALEURAXE ';

  ChercheCodeSqlChpAxeLib(nivMax,TabCodeAxe,chpArbreLib,joinArbreLib);

  if UserLimit <> '' then
  begin
    for i:=1 to nivMax-1 do
    begin
      Values := ReadTokenPipe(UserLimit,'@@');
      if ((values <> '("")') AND (values <> '')) then SQLUserLimit:=SQLUserLimit + ' AND QBR_VALAXENIV'+IntToStr(i)+' IN '+Values;
    end
  end;

  Case ContextBP of
    0,1,2 : begin
            if (sTypeDA='TYPEDA') then
            begin
              MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+
                          codeChpSql+',QBQ_DATECT,QBQ_CODEMAILLE,'+
                          'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
                          'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
                          'QBQ_CAPREAL2,QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
                          'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
                          'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
                          'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
                          'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6) '+
                          ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
                          codesqlArbre+chpArbreLib+',QBR_DATEDELAI,'+STRFPOINT(Mi.code)+
                          ',0,QBR_OP1,0,'+
                          '0,QBR_QTEC,0,'+
                          '0,0,QBR_OP2,0,'+
                          '0,QBR_OP3,0,'+
                          '0,QBR_OP4,0,'+
                          '0,QBR_OP5,0, '+
                          '0,QBR_OP6,0 '+
                          ' FROM QBPARBRE '+joinArbreLib+
                          ' WHERE QBR_CODESESSION="'+codeSession+
                          '" AND QBR_NIVEAU="'+IntToStr(NivMax)+
                          '" '+codeWhere,
                          'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
            end
            else
            begin
              { EVI / Ajout du champ QBQ_UTILISATEUR }
              { EVI / Ajout de QBQ_DATEJOUR = QBH_DATEPIECE aka GL_DATEPIECE }
              //Appel des valeurs de la table QBPARBREDETAIL
              MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,'+
                          codeChpSql+',QBQ_DATECT,QBQ_CODEMAILLE,'+
                          'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
                          'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO,'+
                          'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
                          'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
                          'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
                          'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
                          'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
                          'QBQ_UTILISATEUR,QBQ_DATEJOUR) '+
                          ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
                          codesqlArbre+chpArbreLib+',QBH_DATEDELAI,'+STRFPOINT(Mi.code)+
                          ',0,QBH_OP1,0,'+
                          '0,QBH_QTEC,0,'+
                          '0,QBH_OP2,0,'+
                          '0,QBH_OP3,0,'+
                          '0,QBH_OP4,0,'+
                          '0,QBH_OP5,0,'+
                          '0,QBH_OP6,0,"'
                          +V_PGI.User+
                          '",QBH_DATEPIECE FROM QBPARBREDETAIL LEFT JOIN QBPARBRE ON (QBR_CODESESSION=QBH_CODESESSION AND QBR_NUMNOEUD=QBH_NUMNOEUDREF)'+
                          joinArbreLib+' WHERE QBH_CODESESSION="'+codeSession+
                          '" AND QBH_NIVEAU="'+IntToStr(NivMax)+
                          '" '+codeWhere+SQLUserLimit,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
            end;
          end;
    3 : begin
          //*************** Insertion du PREVU
          { EVI / Ajout du champ QBQ_UTILISATEUR }
          { EVI / Ajout de QBQ_DATEJOUR = QBH_DATEPIECE aka GL_DATEPIECE }
          //Appel des valeurs de la table QBPARBREDETAIL
          MExecuteSql('INSERT INTO ' + GetBase (BaseSQL, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,'+
                      codeChpSql+',QBQ_DATECT,QBQ_CODEMAILLE,'+
                      'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                      'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                      'QBQ_VALAFF6,QBQ_VALAFF7,'+
                      'QBQ_UTILISATEUR,QBQ_DATEJOUR) '+
                      ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
                      codesqlArbre+chpArbreLib+',QBH_DATEDELAI,'+STRFPOINT(Mi.code)+',"Budget N"'+
                      ',QBH_OP1,QBH_QTEC,QBH_OP2,QBH_OP3,QBH_OP4,QBH_OP5,QBH_OP6,"'
                      +V_PGI.User+
                      '",QBH_DATEPIECE FROM ' + GetBase (BaseSQL, 'QBPARBREDETAIL') + ' LEFT JOIN QBPARBRE ON (QBR_CODESESSION=QBH_CODESESSION AND QBR_NUMNOEUD=QBH_NUMNOEUDREF)'+
                      joinArbreLib+' WHERE QBH_CODESESSION="'+codeSession+
                      '" AND QBH_NIVEAU="'+IntToStr(NivMax)+
                      '" '+codeWhere+SQLUserLimit,
                      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');

          //*************** Insertion de l'HISTORIQUE THEORIQUE
          MExecuteSql('INSERT INTO ' + GetBase (BaseSQL, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,'+
                      codeChpSql+',QBQ_DATECT,QBQ_CODEMAILLE,'+
                      'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                      'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                      'QBQ_VALAFF6,QBQ_VALAFF7,'+
                      'QBQ_UTILISATEUR,QBQ_DATEJOUR) '+
                      ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'",'+
                      codesqlArbre+chpArbreLib+',QBH_DATEDELAI,'+STRFPOINT(Mi.code)+',"Théorique"'+
                      ',QBH_REF1,QBH_QTEREF,QBH_REF2,QBH_REF3,QBH_REF4,QBH_REF5,QBH_REF6,"'
                      +V_PGI.User+
                      '",QBH_DATEPIECE FROM ' + GetBase (BaseSQL, 'QBPARBREDETAIL') + ' LEFT JOIN QBPARBRE ON (QBR_CODESESSION=QBH_CODESESSION AND QBR_NUMNOEUD=QBH_NUMNOEUDREF)'+
                      joinArbreLib+' WHERE QBH_CODESESSION="'+codeSession+
                      '" AND QBH_NIVEAU="'+IntToStr(NivMax)+
                      '" '+codeWhere+SQLUserLimit,
                      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
        end;
  end; //CASE
end;

{ EVI / Ajout Date Edition Debut, Date Edition Fin }
procedure RemplitTableQBPCubeTmpPgi0(
          codeSession,BaseSession,codeChpSql,champs,ChpTotal,joinChpSql,
          codeRestriction,UserLimit,SQLUserLimit:hString;
          TabCodeAxe:array of hString;
          DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin:TDateTime);
var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    ListMaille:TListMaille;
    i,NbValAff,NivMax:integer;
    Mi:TMaille;
    types,whereH,whereR,join,DateSQLReq,DateSQLGrpReq,RestrictArticle:hString;
    PrefixH,PrefixR,champsReq,joinChpSqlReq,codeRestrictionReq,SQLUserLimitReq:string;
    DatesSalarie,ReqNewVal,LeftJoinAxe,where0H,where0R,where0:string;
    TableReq,codeWhereR,codeChpValH,codeChpValR,ChpValReq,ETPReq : String;
    Q : Tquery;
begin
  //COrrection alimentation du WhereCtx
  rempliTablePourContexte();

  //cherche les dates de la session
  ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);

  NivMax := ChercheNivMaxSession(codeSession);
  if SessionDelai(codeSession) then nivMax:=nivMax+1;

  ListMaille:=TListMaille.create();
  types:=SessionBPInitialise(codeSession);
  if types='' then types:='0';

  Case ContextBP of
    0,1 : begin //Mode-GC
          if AxeDeviseOK(codeSession) then
          begin
            codeChpValR:=',SUM(GL_TOTALTTCDEV),0,0,'+
                         'SUM(GL_QTEFACT),0,0,'+
                         'SUM(GL_TOTALHTDEV),0,0,'+
                         'SUM(IIF((GL_PCB=0),0,(GL_PUHTNETDEV*GL_QTEFACT)/GL_PCB)),0,0,'+
                         'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNETDEV*GL_QTEFACT)/GL_PCB)),0,0,'+
                         'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0,'+
                         'SUM(GL_TOTALHTDEV)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0  ';
            codeChpValH:=',0,0,SUM(GL_TOTALTTCDEV),'+
                         '0,0,SUM(GL_QTEFACT),'+
                         '0,0,SUM(GL_TOTALHTDEV),'+
                         '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNETDEV*GL_QTEFACT)/GL_PCB)),'+
                         '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNETDEV*GL_QTEFACT)/GL_PCB)),'+
                         '0,0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
                         '0,0,SUM(GL_TOTALHTDEV)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB))';
          end
          else
          begin
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
          end;
          TableReq:=' LIGNE ';
          join:='';
        end;
    2 : begin //Compta
          if AxeDeviseOK(codeSession) then
          begin
            codeChpValR:=',SUM(Y_DEBITDEV-Y_CREDITDEV),0,0,'+
                         '0,0,0,'+
                         'SUM(Y_CREDITDEV-Y_DEBITDEV),0,0,'+
                         '0,0,0,'+
                         '0,0,0,0,0,0,0,0,0  ';
            codeChpValH:=',0,0,SUM(Y_DEBITDEV-Y_CREDITDEV),'+
                         '0,0,0,'+
                         '0,0,SUM(Y_CREDITDEV-Y_DEBITDEV),'+
                         '0,0,0,'+
                         '0,0,0,0,0,0,0,0,0';
          end
          else
          begin
            codeChpValR:=',SUM(Y_DEBIT-Y_CREDIT),0,0,'+
                         '0,0,0,'+
                         'SUM(Y_CREDIT-Y_DEBIT),0,0,'+
                         '0,0,0,'+
                         '0,0,0,0,0,0,0,0,0  ';
            codeChpValH:=',0,0,SUM(Y_DEBIT-Y_CREDIT),'+
                         '0,0,0,'+
                         '0,0,SUM(Y_CREDIT-Y_DEBIT),'+
                         '0,0,0,'+
                         '0,0,0,0,0,0,0,0,0';
          end;
          TableReq:=' ANALYTIQ ';
          join:='';
        end;
    3 : begin
          ReqValAffCube(codeSession,BaseSession,champs,NivMax-1,NbValAff,TableReq,codeWhereR,codeChpValR);

          join:='';
        end;

  end; //CASE

  // Récupération des coefficients Objectif dans la table QBPARBRE
  { EVI / Ajout QBR_VALMODIF pour bug pourcentage d'évolution = 0 }
  if DonneMethodeSession(codeSession) = '1' then
  begin
    Q:=MOpenSql('SELECT QBR_NUMNOEUD'+
    ' FROM ' + GetBase (BaseSession, 'QBPARBRE') + ' WHERE QBR_CODESESSION="'+codeSession+'" AND'+
    ' (QBR_PRCTVARIATION1<>100 OR QBR_PRCTVARIATION2<>100 OR QBR_PRCTVARIATION3<>100'+
    ' OR QBR_PRCTVARIATION4<>100 OR QBR_PRCTVARIATION5<>100 OR QBR_PRCTVARIATION6<>100'+
    ' OR QBR_PRCTVARIATIONQ<>100 OR QBR_VALMODIF="X") AND QBR_NIVEAU='+IntToSTR(NivMax),'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).',true );

    while not Q.eof do
    begin

      if not MoveCurProgressForm(TraduireMemoire('Mise à jour des coefficients...')) then exit;

      //Insertion des coefficients dans la table QBPARBREDETAIL
      MExecuteSQL('UPDATE ' + GetBase (BaseSession, 'QBPARBREDETAIL') + ''+
      ' SET QBH_PRCTVARIATION1=(SELECT QBR_PRCTVARIATION1 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION2=(SELECT QBR_PRCTVARIATION2 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION3=(SELECT QBR_PRCTVARIATION3 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION4=(SELECT QBR_PRCTVARIATION4 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION5=(SELECT QBR_PRCTVARIATION5 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION6=(SELECT QBR_PRCTVARIATION6 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATIONQ=(SELECT QBR_PRCTVARIATIONQ FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+')'+
      ' WHERE QBH_CODESESSION="'+codeSession+'" AND QBH_NUMNOEUDREF='+Q.fields[0].asString,
      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');

      //Calcul des objectifs dans la table QBPARBREDETAIL
      MExecuteSQL('UPDATE ' + GetBase (BaseSession, 'QBPARBREDETAIL') + ''+
      ' SET QBH_OP1=QBH_REF1 + (QBH_REF1*(QBH_PRCTVARIATION1-100)/100),'+
      ' QBH_OP2=QBH_REF2 + (QBH_REF2*(QBH_PRCTVARIATION2-100)/100),'+
      ' QBH_OP3=QBH_REF3 + (QBH_REF3*(QBH_PRCTVARIATION3-100)/100),'+
      ' QBH_OP4=QBH_REF4 + (QBH_REF4*(QBH_PRCTVARIATION4-100)/100),'+
      ' QBH_OP5=QBH_REF5 + (QBH_REF5*(QBH_PRCTVARIATION5-100)/100),'+
      ' QBH_OP6=QBH_REF6 + (QBH_REF6*(QBH_PRCTVARIATION6-100)/100),'+
      ' QBH_QTEC=QBH_QTEREF + (QBH_QTEREF*(QBH_PRCTVARIATIONQ-100)/100)'+
      ' WHERE QBH_CODESESSION="'+codeSession+'" AND QBH_NUMNOEUDREF='+Q.fields[0].asString,
      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
      Q.next;
    end;
    ferme(Q);

    // Mise à jour des axes sans historique
    Q:=MOpenSql('SELECT QBR_NUMNOEUD'+
                ' FROM ' + GetBase (BaseSession, 'QBPARBRE') + ' WHERE QBR_CODESESSION="'+codeSession+'" AND'+
                ' QBR_REF1=0 AND QBR_REF2=0 AND QBR_REF3=0 AND QBR_REF4=0'+
                ' AND QBR_REF5=0 AND QBR_REF6=0 AND QBR_QTEREF=0'+
                ' AND QBR_NIVEAU='+IntToSTR(NivMax),'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).',true );
    while not Q.eof do
    begin
      if not MoveCurProgressForm(TraduireMemoire('Mise à jour des coefficients...')) then exit;
      MExecuteSQL('UPDATE ' + GetBase (BaseSession, 'QBPARBREDETAIL') + ''+
      ' SET QBH_OP1=(SELECT QBR_OP1 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_OP2=(SELECT QBR_OP2 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_OP3=(SELECT QBR_OP3 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_OP4=(SELECT QBR_OP4 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_OP5=(SELECT QBR_OP5 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_OP6=(SELECT QBR_OP6 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_QTEC=(SELECT QBR_QTEC FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+')'+
      ' WHERE QBH_CODESESSION="'+codeSession+'" AND QBH_NUMNOEUDREF='+Q.fields[0].asString,
      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
      Q.next;
    end;
    ferme(Q);
  end;

  //Dans le cas d'une session en Loi d'Eclatement, on initialise les mailles en prenant
  //pour délai minimal le délai minimum retenu pour l'ensemble des lois de la session.
  if SessionEclateeParDelai(codeSession) then types:=SessionMinLoi(codeSession);

  // si on veut le realise jour à jour types:='1';

  InitialiseListeMaille(VALEURI(types),DateDebC,DateFinC,DateDebRef,DateFinRef,ListMaille);

  for i:=0 to ListMaille.count-1 do
  begin

    if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
    begin
      PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), 'Session : '+codeSession);
      break;
    end;

    Mi:=TMaille(ListMaille[i]);

    ////////////////////////////////////////////////////////////////////////////////////////
    // MODIF LM 08/02/2006 pour SPAG
    // Ajouter une clause pour ne prendre en compte que les articles marchandises+prestations
    /////////////////////////////////////////////////////////////////////////////////////////

    Case ContextBP of
      0,1 : begin //Mode-GC
            WhereH:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebReference)+
                    '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinReference)+'" ';
            WhereR:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebCourante)+
                    '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinCourante)+'" ';
            DateSQLReq := ',GL_DATEPIECE';
            DateSQLGrpReq := ',GL_DATEPIECE';
            RestrictArticle := ' AND  GL_TYPELIGNE="ART" AND GL_TYPEARTICLE IN ("MAR", "PRE", "NOM") ';
            champsReq := AnsiReplaceText(champs,'[PREFIXE]','');
            joinChpSqlReq := AnsiReplaceText(joinChpSql,'[PREFIXE]','');
            codeRestrictionReq := AnsiReplaceText(codeRestriction,'[PREFIXE]','');
            SQLUserLimitReq := AnsiReplaceText(SQLUserLimit,'[PREFIXE]','');
          end;
      2 : begin //Compta
            WhereH:=' Y_DATECOMPTABLE>="'+USDATETIME(Mi.DateDebReference)+
                    '" AND Y_DATECOMPTABLE<="'+USDATETIME(Mi.DateFinReference)+'" ';
            WhereR:=' Y_DATECOMPTABLE>="'+USDATETIME(Mi.DateDebCourante)+
                    '" AND Y_DATECOMPTABLE<="'+USDATETIME(Mi.DateFinCourante)+'" ';
            DateSQLReq := ',Y_DATECOMPTABLE';
            DateSQLGrpReq := ',Y_DATECOMPTABLE';
            RestrictArticle := '';
            champsReq := AnsiReplaceText(champs,'[PREFIXE]','');
            joinChpSqlReq := AnsiReplaceText(joinChpSql,'[PREFIXE]','');
            codeRestrictionReq := AnsiReplaceText(codeRestriction,'[PREFIXE]','');
            SQLUserLimitReq := AnsiReplaceText(SQLUserLimit,'[PREFIXE]','');
          end;
      3 : begin //Paie
            where0:='(([PREFIXE]_DATEDEBUT<"[DATEDEBUT]" OR [PREFIXE]_DATEDEBUT>"[DATEFIN]") AND ([PREFIXE]_DATEFIN<"[DATEDEBUT]" OR [PREFIXE]_DATEFIN>"[DATEFIN]"))';
            DatesSalarie := 'PSA_DATEENTREE<="'+USDATETIME(DateFinC)+'" AND (PSA_DATESORTIE>="'+USDATETIME(DateDebRef)+'" OR PSA_DATESORTIE="'+USDateTime(iDate1900)+'")';
            whereH := AnsiReplaceText(codeWhereR,'[DATEDEBUT]',USDATETIME(Mi.DateFinReference));      //CORRECTION SIC DATEDEBUT -> DATEFIN
            whereH := AnsiReplaceText(whereH,'[DATEFIN]',USDATETIME(Mi.DateFinReference));
            whereR := AnsiReplaceText(codeWhereR,'[DATEDEBUT]',USDATETIME(Mi.DateFinCourante));        //CORRECTION SIC DATEDEBUT -> DATEFIN
            whereR := AnsiReplaceText(whereR,'[DATEFIN]',USDATETIME(Mi.DateFinCourante));

            where0H := AnsiReplaceText(where0,'[DATEDEBUT]',USDATETIME(Mi.DateDebReference));
            where0H := AnsiReplaceText(where0H,'[DATEFIN]',USDATETIME(Mi.DateFinReference));
            where0R := AnsiReplaceText(where0,'[DATEDEBUT]',USDATETIME(Mi.DateDebCourante));
            where0R := AnsiReplaceText(where0R,'[DATEFIN]',USDATETIME(Mi.DateFinCourante));

            ChpValReq := AnsiReplaceText(codeChpValR,'[DATEDEBUT]',USDATETIME(Mi.DateDebCourante));
            ChpValReq := AnsiReplaceText(ChpValReq,'[DATEFIN]',USDATETIME(Mi.DateFinCourante));

            DonnePrefixe(codeSession,PrefixH,PrefixR);
            if ((PrefixR = 'NBS') OR (PrefixR = 'EFM') OR (PrefixR = 'ETP') OR (PrefixR = 'PSA')) then PrefixR := 'PPU';
            DateSQLReq := ','+PrefixR+'_DATEDEBUT';
            DateSQLGrpReq := ','+PrefixR+'_DATEDEBUT,'+PrefixR+'_DATEFIN';

            champs := AnsiReplaceText(champs,'[PREFIXE]',PrefixR);

            champsReq := champs;
            champsReq := AnsiReplaceText(champsReq,'TRIM','');
            champsReq := AnsiReplaceText(champsReq,'(','');
            champsReq := AnsiReplaceText(champsReq,')','');

            if pos('ETB_HORAIREETABL',ChpValReq)>0  then ETPReq:=',ETB_HORAIREETABL'
            else ETPReq:='';

            where0H := AnsiReplaceText(where0H,'[PREFIXE]',PrefixR);
            where0R := AnsiReplaceText(where0R,'[PREFIXE]',PrefixR);

            if ((pos('ETB_HORAIREETABL',ChpValReq)>0) AND (pos('ETABCOMPL',joinChpSql)=0)) then
            joinChpSql:=joinChpSql+' LEFT JOIN ETABCOMPL ON [PREFIXE]_ETABLISSEMENT=ETB_ETABLISSEMENT';

            joinChpSqlReq := AnsiReplaceText(joinChpSql,'[PREFIXE]',PrefixR);
            codeRestrictionReq := AnsiReplaceText(codeRestriction,'[PREFIXE]',PrefixR);
            SQLUserLimitReq := AnsiReplaceText(SQLUserLimit,'[PREFIXE]',PrefixR);
            RestrictArticle := '';



          end;
    end;

    { EVI / Remplissage de la table QBPCUBETMP pour les dates d'edition uniquement }
    { EVI / Ajout du champ QBQ_UTILISATEUR }
    if ((DateEdDeb = 0) and (DateEdFin = 0)) or ((Mi.DateDebCourante >= DateEdDeb) and (Mi.DateDebCourante <= DateEdFin)) then
    begin
      Case ContextBP of
        0,1,2 : begin
                //*************** insertion réalisé
                MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                           codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                           'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
                           'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
                           'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
                           'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
                           'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
                           'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
                           'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
                           'QBQ_UTILISATEUR) '+
                           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                           '","'+STRFPOINT(Mi.code)+'",'+champsReq+',"'+USDATETIME(Mi.DateDebCourante)+
                           '"'+DateSQLReq+codeChpValR+',"'+V_PGI.User+
                           '" FROM '+TableReq+joinChpSqlReq+' '+join+
                           ' WHERE '+whereR+codeRestrictionReq+RestrictArticle+SQLUserLimitReq+
                           ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                           'BPCubeTmp (RemplitTableQBPCubeTmpPgiDelaiMaille).');

                //*************** insertion histo
                MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                           'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
                           'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
                           'QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
                           'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
                           'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
                           'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
                           'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
                           'QBQ_UTILISATEUR) '+
                           ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'","'+
                           STRFPOINT(Mi.code)+'",'+champsReq+
                           ',"'+USDATETIME(Mi.DateDebCourante)+'"'+DateSQLReq+CodeChpValH+',"'+V_PGI.User+
                           '" FROM '+TableReq+joinChpSqlReq+' '+join+
                           ' WHERE '+whereH+codeRestrictionReq+RestrictArticle+SQLUserLimitReq+
                           ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                           'BPCubeTmp (RemplitTableQBPCubeTmpPgiDelaiMaille).');
              end;

        3 : begin
              //*************** Insertion REALISE N
             MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '"'+DateSQLReq+',"Réalisé N"'+ChpValReq+',"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+whereR+' AND ' +DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');

              //*************** Insertion REALISE N-1
              MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '"'+DateSQLReq+',"Réalisé N-1"'+ChpValReq+',"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+whereH+' AND ' +DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');

              //*************** Insertion REALISE N - Valeurs nulles
              MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT DISTINCT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '","'+USDATETIME(Mi.DateDebCourante)+'","Réalisé N",0,0,0,0,0,0,0,"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');


              //*************** Insertion HISTORIQUE - Valeurs nulles
              MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT DISTINCT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '","'+USDATETIME(Mi.DateDebCourante)+'","Théorique",0,0,0,0,0,0,0,"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');

              //*************** Insertion PREVU - Valeurs nulles
              MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT DISTINCT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '","'+USDATETIME(Mi.DateDebCourante)+'","Budget N",0,0,0,0,0,0,0,"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+ETPReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');

              //*************** Insertion REALISE N-1 - Valeurs nulles
              MExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
                          codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
                         'QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,'+
                         'QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                         'QBQ_VALAFF6,QBQ_VALAFF7,'+
                         'QBQ_UTILISATEUR) '+
                         ' SELECT DISTINCT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
                         '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
                         '","'+USDATETIME(Mi.DateDebReference)+'","Réalisé N-1",0,0,0,0,0,0,0,"'+V_PGI.User+
                         '" FROM '+TableReq+joinChpSqlReq+' '+join+
                         ' WHERE '+DatesSalarie+codeRestrictionReq+RestrictArticle+
                         ' GROUP BY '+champsReq+DateSQLGrpReq,
                         'BPCubeTmp (RemplitTableQBPCubeTmpPgi0).');

            end;
      end; //CASE

    //**************** INSERTION PREVU
    RemplitTableQBPCubeTmpPrevuObjDelaiMaille(codeSession,BaseSession,codeChpSql,UserLimit,
                                      Mi,TabCodeAxe);

    end;
  end;

  if ContextBP=3 then
  begin
    Q:=OpenSql('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PTK"',true);
    while not Q.eof do
    begin
      if ReqNewVal ='' then ReqNewVal :=  'QBQ_VALAXECT'+IntToStr(NivMax-1)+' LIKE "'+Q.fields[0].asString+'%"'
      else ReqNewVal :=  ReqNewVal + ' OR QBQ_VALAXECT'+IntToStr(NivMax-1)+' LIKE "'+Q.fields[0].asString+'%"';
      Q.next;
    end;
    ferme(Q);

    for i:=0 to ListMaille.count-1 do
    begin

      if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
      begin
        PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), 'Session : '+codeSession);
        break;
      end;

      Mi:=TMaille(ListMaille[i]);

      if ((DateEdDeb = 0) and (DateEdFin = 0)) or ((Mi.DateDebCourante >= DateEdDeb) and (Mi.DateDebCourante <= DateEdFin)) then
      begin
        //*************** Insertion Nouvelles embauches - Valeurs nulles
        ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                   'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE) '+
                   'SELECT DISTINCT QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,"Théorique",0,0,0,0,0,0,0,"'+USDATETIME(Mi.DateDebCourante)+
                   '","'+STRFPOINT(Mi.code)+'" FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' WHERE QBQ_CODESESSION="'+codeSession+'" AND '+
                   '('+ReqNewVal+')');

        ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                   'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE) '+
                   'SELECT DISTINCT QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,"Budget N",0,0,0,0,0,0,0,"'+USDATETIME(Mi.DateDebCourante)+
                   '","'+STRFPOINT(Mi.code)+'" FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' WHERE QBQ_CODESESSION="'+codeSession+'" AND '+
                   '('+ReqNewVal+')');

        ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                   'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE) '+
                   'SELECT DISTINCT QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,"Réalisé N",0,0,0,0,0,0,0,"'+USDATETIME(Mi.DateDebCourante)+
                   '","'+STRFPOINT(Mi.code)+'" FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' WHERE QBQ_CODESESSION="'+codeSession+'" AND '+
                   '('+ReqNewVal+')');

        ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
                   'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE) '+
                   'SELECT DISTINCT QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
                   'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
                   'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
                   'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
                   'QBQ_CTX,QBQ_UTILISATEUR,"Réalisé N-1",0,0,0,0,0,0,0,"'+USDATETIME(Mi.DateDebCourante)+
                   '","'+STRFPOINT(Mi.code)+'" FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' WHERE QBQ_CODESESSION="'+codeSession+'" AND '+
                   '('+ReqNewVal+')');
      end;
    end;


    For i:=1 to NivMax-1 do
      LeftJoinAxe:= LeftJoinAxe + ' AND T1.QBQ_VALAXECT'+IntToStr(i)+'=T2.QBQ_VALAXECT'+IntToStr(i);

    //Insertion Ecart Mensuel N N-1
    ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
               'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
               'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
               'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
               'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
               'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE)'+
               'SELECT DISTINCT T1.QBQ_VALAXECT1,T1.QBQ_VALAXECT2,T1.QBQ_VALAXECT3,T1.QBQ_VALAXECT4,T1.QBQ_VALAXECT5,'+
               'T1.QBQ_VALAXECT6,T1.QBQ_VALAXECT7,T1.QBQ_VALAXECT8,T1.QBQ_VALAXECT9,T1.QBQ_VALAXECT10,T1.QBQ_LIBVALAXECT1,'+
               'T1.QBQ_LIBVALAXECT2,T1.QBQ_LIBVALAXECT3,T1.QBQ_LIBVALAXECT4,T1.QBQ_LIBVALAXECT5,T1.QBQ_LIBVALAXECT6,'+
               'T1.QBQ_LIBVALAXECT7,T1.QBQ_LIBVALAXECT8,T1.QBQ_LIBVALAXECT9,T1.QBQ_LIBVALAXECT10,T1.QBQ_CODESESSION,'+
               'T1.QBQ_CTX,T1.QBQ_UTILISATEUR,"N vs N-1",SUM(T2.QBQ_VALAFF1-T1.QBQ_VALAFF1),SUM(T2.QBQ_VALAFF2-T1.QBQ_VALAFF2),'+
               'SUM(T2.QBQ_VALAFF3-T1.QBQ_VALAFF3),SUM(T2.QBQ_VALAFF4-T1.QBQ_VALAFF4),SUM(T2.QBQ_VALAFF5-T1.QBQ_VALAFF5),'+
               'SUM(T2.QBQ_VALAFF6-T1.QBQ_VALAFF6),SUM(T2.QBQ_VALAFF7-T1.QBQ_VALAFF7),T1.QBQ_DATECT,T1.QBQ_CODEMAILLE '+
               'FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' AS T1 LEFT JOIN ' + GetBase (BaseSession, 'QBPCUBETMP') + ' AS T2 ON T1.QBQ_VALAXECT1=T2.QBQ_VALAXECT1 '+LeftJoinAxe+
               ' AND T2.QBQ_NATVALAFF="Réalisé N" AND T1.QBQ_CODESESSION=T2.QBQ_CODESESSION AND T1.QBQ_CODEMAILLE=T2.QBQ_CODEMAILLE '+
               'WHERE T1.QBQ_CODESESSION="'+codeSession+'" AND T1.QBQ_NATVALAFF="Réalisé N-1" GROUP BY T1.QBQ_VALAXECT1,T1.QBQ_VALAXECT2,'+
               'T1.QBQ_VALAXECT3,T1.QBQ_VALAXECT4,T1.QBQ_VALAXECT5,T1.QBQ_VALAXECT6,T1.QBQ_VALAXECT7,T1.QBQ_VALAXECT8,'+
               'T1.QBQ_VALAXECT9,T1.QBQ_VALAXECT10,T1.QBQ_LIBVALAXECT1,T1.QBQ_LIBVALAXECT2,T1.QBQ_LIBVALAXECT3,T1.QBQ_LIBVALAXECT4,'+
               'T1.QBQ_LIBVALAXECT5,T1.QBQ_LIBVALAXECT6,T1.QBQ_LIBVALAXECT7,T1.QBQ_LIBVALAXECT8,T1.QBQ_LIBVALAXECT9,'+
               'T1.QBQ_LIBVALAXECT10,T1.QBQ_CODESESSION,T1.QBQ_CTX,T1.QBQ_UTILISATEUR,T1.QBQ_DATECT,T1.QBQ_CODEMAILLE');

    //Insertion Ecart Mensuel Réalisé - Prévu
    ExecuteSql('INSERT INTO ' + GetBase (BaseSession, 'QBPCUBETMP') + ' (QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,'+
               'QBQ_VALAXECT6,QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10,QBQ_LIBVALAXECT1,'+
               'QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
               'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_CODESESSION,'+
               'QBQ_CTX,QBQ_UTILISATEUR,QBQ_NATVALAFF,QBQ_VALAFF1,QBQ_VALAFF2,QBQ_VALAFF3,QBQ_VALAFF4,QBQ_VALAFF5,'+
               'QBQ_VALAFF6,QBQ_VALAFF7,QBQ_DATECT,QBQ_CODEMAILLE)'+
               'SELECT DISTINCT T1.QBQ_VALAXECT1,T1.QBQ_VALAXECT2,T1.QBQ_VALAXECT3,T1.QBQ_VALAXECT4,T1.QBQ_VALAXECT5,'+
               'T1.QBQ_VALAXECT6,T1.QBQ_VALAXECT7,T1.QBQ_VALAXECT8,T1.QBQ_VALAXECT9,T1.QBQ_VALAXECT10,T1.QBQ_LIBVALAXECT1,'+
               'T1.QBQ_LIBVALAXECT2,T1.QBQ_LIBVALAXECT3,T1.QBQ_LIBVALAXECT4,T1.QBQ_LIBVALAXECT5,T1.QBQ_LIBVALAXECT6,'+
               'T1.QBQ_LIBVALAXECT7,T1.QBQ_LIBVALAXECT8,T1.QBQ_LIBVALAXECT9,T1.QBQ_LIBVALAXECT10,T1.QBQ_CODESESSION,'+
               'T1.QBQ_CTX,T1.QBQ_UTILISATEUR,"Réalisé N vs Budget N",SUM(T2.QBQ_VALAFF1-T1.QBQ_VALAFF1),SUM(T2.QBQ_VALAFF2-T1.QBQ_VALAFF2),'+
               'SUM(T2.QBQ_VALAFF3-T1.QBQ_VALAFF3),SUM(T2.QBQ_VALAFF4-T1.QBQ_VALAFF4),SUM(T2.QBQ_VALAFF5-T1.QBQ_VALAFF5),'+
               'SUM(T2.QBQ_VALAFF6-T1.QBQ_VALAFF6),SUM(T2.QBQ_VALAFF7-T1.QBQ_VALAFF7),T1.QBQ_DATECT,T1.QBQ_CODEMAILLE '+
               'FROM ' + GetBase (BaseSession, 'QBPCUBETMP') + ' AS T1 LEFT JOIN ' + GetBase (BaseSession, 'QBPCUBETMP') + ' AS T2 ON T1.QBQ_VALAXECT1=T2.QBQ_VALAXECT1 '+LeftJoinAxe+
               ' AND T2.QBQ_NATVALAFF="Réalisé N" AND T1.QBQ_CODESESSION=T2.QBQ_CODESESSION AND T1.QBQ_CODEMAILLE=T2.QBQ_CODEMAILLE '+
               'WHERE T1.QBQ_CODESESSION="'+codeSession+'" AND T1.QBQ_NATVALAFF="Budget N" GROUP BY T1.QBQ_VALAXECT1,T1.QBQ_VALAXECT2,'+
               'T1.QBQ_VALAXECT3,T1.QBQ_VALAXECT4,T1.QBQ_VALAXECT5,T1.QBQ_VALAXECT6,T1.QBQ_VALAXECT7,T1.QBQ_VALAXECT8,'+
               'T1.QBQ_VALAXECT9,T1.QBQ_VALAXECT10,T1.QBQ_LIBVALAXECT1,T1.QBQ_LIBVALAXECT2,T1.QBQ_LIBVALAXECT3,T1.QBQ_LIBVALAXECT4,'+
               'T1.QBQ_LIBVALAXECT5,T1.QBQ_LIBVALAXECT6,T1.QBQ_LIBVALAXECT7,T1.QBQ_LIBVALAXECT8,T1.QBQ_LIBVALAXECT9,'+
               'T1.QBQ_LIBVALAXECT10,T1.QBQ_CODESESSION,T1.QBQ_CTX,T1.QBQ_UTILISATEUR,T1.QBQ_DATECT,T1.QBQ_CODEMAILLE');
  end;

  freeListMaille(ListMaille);

end;

{ EVI / Ajout Date Edition Debut, Date Edition Fin }
procedure RemplitTableQBPCubeTmpPgi(const codeSession,BaseSQL,UserLimit:hString;
          DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin : TDateTime);
var
  codeChpSql,ChpTotal,codeRestriction,champs,joinChpSql,SQLUserLimit:hString;
  chpLib,joinLib,codeJoinR:hString;
  nivMax,i, j:integer;
  TabCodeAxe,TabNumValAxe:array [1..11] of hString;
  TS, TS2, TS3 : TstringList;
  s, LJSQL, Alias, AliasOk, AliasSup : hString;
  {$IFNDEF MODE}{$IFDEF GCGC} sSessionN, sSessionN1 : hString; {$ENDIF}{$ENDIF}  //CEGID-CCMX le 06/11/2006
begin
  //vide la table temporaire
  { EVI / Ajout du champ QBQ_UTILISATEUR }
  MExecuteSql('DELETE FROM ' + GetBase (BaseSQL, 'QBPCUBETMP') + ' WHERE QBQ_CODESESSION="'+codeSession+'" AND QBQ_UTILISATEUR="'+V_PGI.User +'"',
              'BPCubeTmp (RemplitTableQBPCubeTmpPgi).');

  { EVI / Gestion des sessions en Loi d'Eclatement
  Recopie de la table QBPARBRE dans la table QBPARBREDETAIL avant génération du cube }
  if DonneMethodeSession(codeSession) = '2' then
  begin
    MExecuteSql('DELETE FROM ' + GetBase (BaseSQL, 'QBPARBREDETAIL') +' WHERE QBH_CODESESSION="'+codeSession+'"',
                'BPCubeTMP (RemplitTableQBPCubeTmpPgi).');


    MExecuteSql('INSERT INTO ' + GetBase (BaseSQL, 'QBPARBREDETAIL') +' (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
                'QBH_DATEDELAI,'+
                'QBH_OP1,QBH_OPPRCT1,QBH_REF1,QBH_REFPRCT1,'+
                'QBH_OP2,QBH_OPPRCT2,QBH_REF2,QBH_REFPRCT2,'+
                'QBH_OP3,QBH_OPPRCT3,QBH_REF3,QBH_REFPRCT3,'+
                'QBH_OP4,QBH_OPPRCT4,QBH_REF4,QBH_REFPRCT4,'+
                'QBH_OP5,QBH_OPPRCT5,QBH_REF5,QBH_REFPRCT5,'+
                'QBH_OP6,QBH_OPPRCT6,QBH_REF6,QBH_REFPRCT6,'+
                'QBH_QTEC,QBH_QTECPRCT,QBH_QTEREF,QBH_QTEREFPRCT,'+
                'QBH_PRCTVARIATION1,QBH_PRCTVARIATION2,QBH_PRCTVARIATION3,'+
                'QBH_PRCTVARIATION4,QBH_PRCTVARIATION5,QBH_PRCTVARIATION6,'+
                'QBH_PRCTVARIATIONQ,QBH_DATEPIECE,QBH_NUMNOEUDREF)'+
                ' SELECT QBR_CODESESSION,QBR_NUMNOEUD,QBR_NIVEAU,'+
                'QBR_DATEDELAI,'+
                'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                'QBR_PRCTVARIATION1,QBR_PRCTVARIATION2,QBR_PRCTVARIATION3,'+
                'QBR_PRCTVARIATION4,QBR_PRCTVARIATION5,QBR_PRCTVARIATION6,'+
                'QBR_PRCTVARIATIONQ,QBR_DATEDELAI,QBR_NUMNOEUD'+
                ' FROM ' + GetBase (BaseSQL, 'QBPARBRE') +' WHERE QBR_CODESESSION="'+codeSession+'"','BPCubeTMP (RemplitTableQBPCubeTmpPgi).');

  end;

  ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

  nivMax:=ChercheNivMax(codeSession);

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
 ChercheCodeSqlChpAxeLibPGI(nivMax,false,TabCodeAxe,'',chpLib,joinLib);
 codeRestriction:=ChercheCodeRestrictionSession('','','',codeSession,codeJoinR);
 if UserLimit <> '' then DonneCodeChpUserRestrict(TabCodeAxe,UserLimit, nivMax, SQLUserLimit);

 LJSQL := '';
 TS := TStringList.Create;
 try
   TS2 := TStringList.Create;
   try
     TS3 := TStringList.Create;
     try
       LJSQL := AnalyseJoinSQL(joinChpSql+joinLib+CodeJoinR, TS, TS2, TS3);
       //Suppression des alias des champs
       for j := 0 to TS2.Count - 1 do
       begin
         s := TS2[j];
         Alias   := ReadTokenSt(s);
         AliasOk := ReadTokenSt(s);
         if not StrToBool_(AliasOk) then
         begin
           AliasSup := Alias + ';' + AliasSup;
           if ((TS3[j] <> '') AND (Pos(TS3[j],AliasSup)=0)) then
           begin
             codeChpSql      := StringReplace(codeChpSql     , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             champs          := StringReplace(champs         , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             chpLib          := StringReplace(chpLib         , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             ChpTotal        := StringReplace(ChpTotal       , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             codeRestriction := StringReplace(codeRestriction, Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             LJSQL           := StringReplace(LJSQL          , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
             SQLUserLimit    := StringReplace(SQLUserLimit  , Alias+'.', TS3[j]+'.', [rfReplaceAll]);
           end
           else
           begin
             codeChpSql      := StringReplace(codeChpSql     , Alias+'.', '', [rfReplaceAll]);
             champs          := StringReplace(champs         , Alias+'.', '', [rfReplaceAll]);
             chpLib          := StringReplace(chpLib         , Alias+'.', '', [rfReplaceAll]);
             ChpTotal        := StringReplace(ChpTotal       , Alias+'.', '', [rfReplaceAll]);
             codeRestriction := StringReplace(codeRestriction, Alias+'.', '', [rfReplaceAll]);
             LJSQL           := StringReplace(LJSQL          , Alias+'.', '', [rfReplaceAll]);
             SQLUserLimit    := StringReplace(SQLUserLimit   , Alias+'.', '', [rfReplaceAll]);
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

 //CEGID-CCMX le 07/11/2006 DEBUT : verification si session liée aux DA
 // le 17/11/2006 : si session budgetDA de l'année encours ou année suivante
 // le 16/01/2007 : en fait on limite pas seulement au session N ou N+1 mais aux
 // sessions utilisées

  {$IFNDEF MODE}
  {$IFDEF GCGC}
  sSessionN := GetParamsocSecur('SO_GCCODESESSIONN','');
  sSessionN1 := GetParamsocSecur('SO_GCCODESESSIONN1','');
  if (sSessionN = codesession) or (sSessionN1 = codesession)
  or (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + codesession+'"'))
  then
    RemplitTableQBPCubeTmpPgiTYPEDA(codeSession,codeChpSql,champs+chpLib,ChpTotal,
           LJSQL,codeRestriction,UserLimit,SQLUserLimit,TabCodeAxe,
           DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin)
  else
 {$ENDIF}
 {$ENDIF}
 { EVI / Ajout Date Edition Debut, Date Edition Fin }
 RemplitTableQBPCubeTmpPgi0(codeSession,BaseSQL,codeChpSql,champs+chpLib,ChpTotal,
           LJSQL,codeRestriction,UserLimit,SQLUserLimit,TabCodeAxe,
           DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin);
 //CEGID-CCMX le 07/11/2006 FIN

 { EVI / Récupération des libellés pour Nouvelle Embauche en Paie }
 if ContextBP = 3 then
 begin
   for i:=low(TabCodeAxe) to high(TabCodeAxe) do
   begin
     if TabCodeAxe[i]='011' then
     begin
       MajCubeLibelleNewValues(codeSession,i);
       break;
     end;
   end;
 end;

 { EVI / Réinitialise le flag VALMODIF }
 MExecuteSQL('UPDATE ' + GetBase (BaseSQL, 'QBPARBRE') +' SET QBR_VALMODIF="-" WHERE QBR_CODESESSION="'+codeSession+'"',
             'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
end;

//CEGID-CCMX le 07/11/2006 DEBUT Fonction qui remplit la teble tempo avec les info liées au type de DA
procedure RemplitTableQBPCubeTmpPgiTYPEDA(
          codeSession,codeChpSql,champs,ChpTotal,joinChpSql,
          codeRestriction,UserLimit,SQLUserLimit:hstring;
          TabCodeAxe:array of hstring;
          DateDebCourante,DateFinCourante,DateEdDeb,DateEdFin:TDateTime);
var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    ListMaille:TListMaille;
    i,NivMax:integer;
    Mi:TMaille;
    types,codeChpValH,codeChpValR, codeChpValPR,whereH,whereR,wherePR,table,join:string;
    sSavChamps,sChp,sChpPieceDA,joinChpPieceDA : string;
    DateSQL,RestrictArticle : string;
    Q : Tquery;
begin
  //CEGID-CCMX le 16/03/2007 : ajout correction faite en standard
  //COrrection alimentation du WhereCtx
  rempliTablePourContexte();
  ////CEGID-CCMX le 16/03/2007 FIN

 //cherche les dates de la session
 ChercheDateDDateFPeriode(codeSession,DateDebC,DateFinC,DateDebRef,DateFinRef);

 NivMax := ChercheNivMaxSession(codeSession);
 if SessionDelai(codeSession) then nivMax:=nivMax+1;

 ListMaille:=TListMaille.create();
 types:=SessionBPInitialise(codeSession);
 if types='' then types:='0';
 
 //CEGID-CCMX le 23/01/2007 : on fait une somme sur les lignes
 // car certaines lignes peuvent être commandées et par d'autres
 //cela fausse le résultat de faire le calcul sur l'en-tête
 codeChpValPR:=',0,0,0,'+
              '0,0,0,'+
              'SUM(DAL_MONTANTHT),0,0,0,'+
              '0,0,0,'+
              '0,0,0,'+
              '0,0,0,'+
              '0,0,0  ';
 codeChpValR:=',SUM(GL_TOTALTTC),0,0,'+
              'SUM(GL_QTEFACT),0,0,'+
              '0,SUM(GL_TOTALHT),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0,'+
              'SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),0,0  ';
 codeChpValH:=',0,0,SUM(GL_TOTALTTC),'+
              '0,0,SUM(GL_QTEFACT),'+
              '0,0,0,SUM(GL_TOTALHT),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUHTNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PUTTCNET*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB)),'+
              '0,0,SUM(GL_TOTALHT)-SUM(IIF((GL_PCB=0),0,(GL_PMAP*GL_QTEFACT)/GL_PCB))';
 Table:=' LIGNE ';
 join:='';

 joinChpPieceDA := ' ';
 champs := AnsiReplaceText(champs,'[PREFIXE]','');
 sSavChamps :=champs;
 Repeat
   sChp:=ReadTokenPipe(champs,',');
   if sChp <> '' then
   begin
    if (sChpPieceDA) <> '' then sChpPieceDA := sChpPieceDA + ', '
                           else sChpPieceDA := ' ';
    //CEGID-CCMX le 16/03/2007 : le contenu de sChp a changé maintenant il y a TRIM(...)
    if (trim(sChp)='TRIM(GLC_CODESERVICE)') then
    begin
      sChpPieceDA := sChpPieceDA + 'DA_CODESERVICE';
      joinChpPieceDA:=joinChpPieceDA+'LEFT JOIN SERVICES  ON PGS_CODESERVICE=DA_CODESERVICE ';
    end
    else if (trim(sChp)='TRIM(GL_ETABLISSEMENT)') then
    begin
      sChpPieceDA := sChpPieceDA + 'DA_ETABLISSEMENT';
      joinChpPieceDA:=joinChpPieceDA+'LEFT JOIN ETABLISS  ON ET_ETABLISSEMENT=DA_ETABLISSEMENT ';
    end
    else if (trim(sChp)='TRIM(GLC_TYPEDA)') then
    begin
      sChpPieceDA := sChpPieceDA + 'DA_TYPEDA';
      joinChpPieceDA:=joinChpPieceDA+'LEFT JOIN TYPEDA  ON DAT_TYPEDA=DA_TYPEDA ';
    end
    else sChpPieceDA := sChpPieceDA + sChp;
   end;
   until sChp = '';
  champs:=sSavChamps;

  // Récupération des coefficients Objectif dans la table QBPARBRE
  { EVI / Ajout QBR_VALMODIF pour bug pourcentage d'évolution = 0 }
  if DonneMethodeSession(codeSession) = '1' then
  begin
    Q:=MOpenSql('SELECT QBR_NUMNOEUD'+
    ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND'+
    ' (QBR_PRCTVARIATION1<>100 OR QBR_PRCTVARIATION2<>100 OR QBR_PRCTVARIATION3<>100'+
    ' OR QBR_PRCTVARIATION4<>100 OR QBR_PRCTVARIATION5<>100 OR QBR_PRCTVARIATION6<>100'+
    ' OR QBR_PRCTVARIATIONQ<>100 OR QBR_VALMODIF="X") AND QBR_NIVEAU='+IntToSTR(NivMax),'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).',true );

    while not Q.eof do
    begin

      if not MoveCurProgressForm(TraduireMemoire('Mise à jour des coefficients...')) then exit;

      //Insertion des coefficients dans la table QBPARBREDETAIL
      MExecuteSQL('UPDATE QBPARBREDETAIL'+
      ' SET QBH_PRCTVARIATION1=(SELECT QBR_PRCTVARIATION1 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION2=(SELECT QBR_PRCTVARIATION2 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION3=(SELECT QBR_PRCTVARIATION3 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION4=(SELECT QBR_PRCTVARIATION4 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION5=(SELECT QBR_PRCTVARIATION5 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATION6=(SELECT QBR_PRCTVARIATION6 FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+'),'+
      ' QBH_PRCTVARIATIONQ=(SELECT QBR_PRCTVARIATIONQ FROM QBPARBRE'+
      ' WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+Q.fields[0].asString+')'+
      ' WHERE QBH_CODESESSION="'+codeSession+'" AND QBH_NUMNOEUDREF='+Q.fields[0].asString,
      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');

      //Calcul des objectifs dans la table QBPARBREDETAIL
      MExecuteSQL('UPDATE QBPARBREDETAIL'+
      ' SET QBH_OP1=QBH_REF1 + (QBH_REF1*(QBH_PRCTVARIATION1-100)/100),'+
      ' QBH_OP2=QBH_REF2 + (QBH_REF2*(QBH_PRCTVARIATION2-100)/100),'+
      ' QBH_OP3=QBH_REF3 + (QBH_REF3*(QBH_PRCTVARIATION3-100)/100),'+
      ' QBH_OP4=QBH_REF4 + (QBH_REF4*(QBH_PRCTVARIATION4-100)/100),'+
      ' QBH_OP5=QBH_REF5 + (QBH_REF5*(QBH_PRCTVARIATION5-100)/100),'+
      ' QBH_OP6=QBH_REF6 + (QBH_REF6*(QBH_PRCTVARIATION6-100)/100),'+
      ' QBH_QTEC=QBH_QTEREF + (QBH_QTEREF*(QBH_PRCTVARIATIONQ-100)/100)'+
      ' WHERE QBH_CODESESSION="'+codeSession+'" AND QBH_NUMNOEUDREF='+Q.fields[0].asString,
      'BPCubeTmp (RemplitTableQBPCubeTmpPrevuObjDelaiMaille).');
      Q.next;
    end;
    ferme(Q);
  end;

  //Dans le cas d'une session en Loi d'Eclatement, on initialise les mailles en prenant
  //pour délai minimal le délai minimum retenu pour l'ensemble des lois de la session.
  if SessionEclateeParDelai(codeSession) then types:=SessionMinLoi(codeSession);

  InitialiseListeMaille(VALEURI(types),DateDebC,DateFinC,
                       DateDebRef,DateFinRef,ListMaille);

  for i:=0 to ListMaille.count-1 do
  begin

   if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
   begin
     PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), 'Session : '+codeSession);
     break;
   end;

   Mi:=TMaille(ListMaille[i]);
   WhereH:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebReference)+
           '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinReference)+'" ';
   WhereR:=' GL_DATEPIECE>="'+USDATETIME(Mi.DateDebCourante)+
           '" AND GL_DATEPIECE<="'+USDATETIME(Mi.DateFinCourante)+'" ';
   WherePR:=' DA_DATELIVRAISON>="'+USDATETIME(Mi.DateDebCourante)+
           '" AND DA_DATELIVRAISON<="'+USDATETIME(Mi.DateFinCourante)+'" ';
   DateSQL := ',GL_DATEPIECE';
   RestrictArticle := ' AND  GL_TYPELIGNE="ART" AND GL_TYPEARTICLE IN ("MAR", "PRE", "NOM") ';
   joinChpSql := AnsiReplaceText(joinChpSql,'[PREFIXE]','');
   codeRestriction := AnsiReplaceText(codeRestriction,'[PREFIXE]','');
   SQLUserLimit := AnsiReplaceText(SQLUserLimit,'[PREFIXE]','');

   { GC/GRC : EV / Remplissage de la table QBPCUBETMP pour les dates d'edition uniquement }
   { CEGID-CCMX le 16/03/2007 : reprise modif EVI / Ajout du champ QBQ_UTILISATEUR }
   if ((DateEdDeb = 0) and (DateEdFin = 0)) or ((Mi.DateDebCourante >= DateEdDeb) and (Mi.DateDebCourante <= DateEdFin)) then
   begin
     //*************** insertion réalisé
     MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
            codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAPREAL2,QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
            'QBQ_UTILISATEUR) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
            '","'+STRFPOINT(Mi.code)+'",'+champs+',"'+USDATETIME(Mi.DateDebCourante)+
            '" '+DateSQL+codeChpValR+',"'+V_PGI.User+
            '" FROM '+Table+joinChpSql+' '+join+
            ' WHERE '+whereR+codeRestriction+RestrictArticle+SQLUserLimit+
            ' GROUP BY '+champs+DateSQL,
            'BPCubeTmp (RemplitTableQBPCubeTmpPgiTYPEDA).');

     //*************** Insertion Presque réalisé
     // on prend la table PIECEDA
     //CEGID-CCMX le 23/01/2007 : on fait une somme sur les lignes
     MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+
            codeChpSql+',QBQ_DATECT,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAPREAL2,QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
            'QBQ_UTILISATEUR) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+
            '","'+STRFPOINT(Mi.code)+'",'+sChpPieceDA+',"'+USDATETIME(Mi.DateDebCourante)+
            '" '+codeChpValPR+',"'+V_PGI.User+
            '" FROM '+' PIECEDA '+joinChpPieceDA+' '+join+
            ' LEFT JOIN LIGNEDA ON DA_SOUCHE = DAL_SOUCHE AND DA_NUMERO = DAL_NUMERO ' +
            ' WHERE '+wherePR+
            ' AND ((DAL_STATUTDA="30" ) OR (DAL_STATUTDA = "10" AND DA_STATUTDA ="20" AND DA_CTRLBUDGET="X") ) '+
            ' GROUP BY '+sChpPieceDA+',DA_DATELIVRAISON',
            'BPCubeTmp (RemplitTableQBPCubeTmpPgiTYPEDA).');

    //*************** insertion histo
    //on prend dans la table ligne
    MExecuteSql('INSERT INTO QBPCUBETMP (QBQ_CTX,QBQ_CODESESSION,QBQ_CODEMAILLE,'+codeChpSql+',QBQ_DATECT,QBQ_DATEJOUR,'+
            'QBQ_CAREALISE,QBQ_CAPREVU,QBQ_CAHISTO,'+
            'QBQ_REALISE,QBQ_PREVU,QBQ_HISTO, '+
            'QBQ_CAPREAL2,QBQ_CAREALISE2,QBQ_CAPREVU2,QBQ_CAHISTO2,'+
            'QBQ_CAREALISE3,QBQ_CAPREVU3,QBQ_CAHISTO3,'+
            'QBQ_CAREALISE4,QBQ_CAPREVU4,QBQ_CAHISTO4,'+
            'QBQ_CAREALISE5,QBQ_CAPREVU5,QBQ_CAHISTO5,'+
            'QBQ_CAREALISE6,QBQ_CAPREVU6,QBQ_CAHISTO6,'+
            'QBQ_UTILISATEUR) '+
            ' SELECT "'+ValeurContexte('qbpcubetmp')+'","'+codeSession+'","'+
            STRFPOINT(Mi.code)+'",'+champs+
            ',"'+USDATETIME(Mi.DateDebCourante)+'"'+DateSQL+CodeChpValH+',"'+V_PGI.User+
            '" FROM '+Table+joinChpSql+' '+join+
            ' WHERE '+whereH+codeRestriction+RestrictArticle+SQLUserLimit+
            ' GROUP BY '+champs+DateSQL,
            'BPCubeTmp (RemplitTableQBPCubeTmpPgiTYPEDA).');

   //**************** insertion prévu
   RemplitTableQBPCubeTmpPrevuObjDelaiMaille(codeSession,'',codeChpSql,UserLimit,
                                     Mi,TabCodeAxe,'TYPEDA');

   end;
  end;
  freeListMaille(ListMaille);
end;
//CEGID-CCMX le 07/11/2006 FIN
end.
