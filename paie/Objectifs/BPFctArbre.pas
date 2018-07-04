unit BPFctArbre;

interface

uses HEnt1,UGraph;

function BPIncrementeNumNoeud(const codeSession:hString):integer;
function BPIncrementeNumNoeudDetail(const codeSession:hString):integer;

//Mise à jour feuille arbre
procedure MiseAJourNoeudPere(const codeSession:hString);
procedure MiseAJourPrctArbre(const codeSession:hString;okMajPrctRef,OkMajPrctObj:boolean);   
procedure MiseAJourQtePrevueQteRetenue(const codeSession:hString);

procedure MiseAjourNiv(MAJTable:boolean;const codeSession,Noeud,NoeudPere,Niveau:hString;
          nivMaxSession:integer;
          TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
procedure MiseAjourNiveauPrecedent(MAJTable:boolean;codeSession,numNoeud,numnoeudpere:hString;
          niveau,nivMaxSession:integer;TotCAVal1,totQteVal,
          TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
          TotalPrevu,TotalRet,CARetenu:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
procedure MiseAjourNiveauSuivant(MAJTable:boolean;numnoeud,codeSession:hString;
          niveau,nivMaxSession,RI:integer;newCAVal1,newQteVal,
          newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,TotalRet,CARetenu:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);

//Vide Arbre
procedure VideArbre(const codeSession:hString;Update:boolean);
procedure VideArbreNiveau(const codeSession:hString;Niveau:integer);

//Valeur bloquée
procedure MiseAJourValeurBloque(const codeSession:hString;NivMax:integer);
function  NoeudCourantValBloquee(const codeSession,numNoeud,numNoeudPere:hString):boolean;

//Total niveau
procedure ArbreTotalNivRefPere(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
procedure ArbreTotalNivRef(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
procedure ArbreTotalNiv(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
          TotalPrevu,TotalRet,CARetenu:double);
procedure ArbreTotalNivHisto(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);

//Total noeud
procedure ArbreValeurNoeud(const codeSession,noeud:hString;
          var newVal,newQteVal,newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,QteRetenue,CARetenu:double);
procedure ArbreValeurRefNoeud(const codeSession,noeud:hString;
          var Val,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double);
procedure ArbreValeurRefNoeudTotal(const codeSession:hString;
          var Val,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double);

//Total arbre
function ChercheTotalArbre(const codeSession,typeC:hString):double;

//Controle Arbre
procedure ControleNoeudANoeudPere(const codeSession:hString);

//Evloution
procedure EvolutionSurUneValeurAxe(const retour,codeSession,codeaxe,valeuraxe:hString;Onglet:integer);

//MetArbreEnLigne
procedure MetArbreEnLigneNiveau(Niveau:integer;const types,session,ChpsqlR,codeFiltre:hString;
          TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalQte:double;
          TabCodeAxe:array of hString;QTBP:TQTob;okCoeff,okPrev:boolean;nbTabAffiche:integer);
procedure MetArbreEnLigne(NivMaxProfil:integer;const types,session:hString;
          total:double;
          TabCodeAxe:array of hString;QTBP:TQTob);

//Noeud pere
function ChercheNoeudPere(codeSession:hString;NoeudCourant:integer):integer;

//Modification d'un noeud
procedure BPUpdate(const SaisiSvg,CorrectifSvg,
          saivalS,correctifvalS:hString;
          var OkModifSaisi,OkModifCorrectif,okMAJCorrectif:boolean;
          var saival,correctifval:double);

//Mise à jour des prct de variation
procedure MAJPrctVariation(const codeSession:hString);

//
function ValeursFilsSuperieurValeurPereBloq(okPere:boolean;codeSession,noeudPere,noeud,
         codeSqlNewValeurNoeud,codeChpModif:hString;
         TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
         TotalAncCA4,TotalAncCA5,TotalAncCA6:double):boolean;
function OkModifNoeudArbre(codeSession,noeud,noeudPere,codeSqlNewValeurNoeud,codeChpModif:hString;
         niveau:integer;
         OkModifValBloque:boolean):boolean;

procedure CreateSousNiveauAuto(const retour,codeSession,niveau,NoeudPere:hString);
procedure ModifSousNiveauAuto(const retour,codeSession,codeaxe,valeuraxe,niveau,NoeudPere:hString);
procedure InsertNewVal(const retour,codeSession:hString);
procedure EvolutionTwo(const retour,codeSession:hString);

implementation

uses Sysutils,HMsgBox,UTob,ed_tools,StrUtils,HCtrls,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     BPBasic,Uutil,BPFctSession,BPUtil;

//*****************************************************************************
//

//fonction qui retourne le nouveau numéro de noeud
//pour une session donnée
function BPIncrementeNumNoeud(const codeSession:hString):integer;
var Q:TQuery;
begin
 result:=10;
 Q:=MOpenSql('SELECT MAX(QBR_NUMNOEUD)+1 FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+'"',
             'BPFctArbre (BPIncrementenumNoeud).',true);
 if not Q.eof
  then
   begin
    if Q.fields[0].asString<>''
     then result:=Q.fields[0].asInteger;
   end;
 ferme(Q);
end;

//fonction qui retourne le nouveau numéro de noeud
//pour une session donnée
function BPIncrementeNumNoeudDetail(const codeSession:hString):integer;
var Q:TQuery;
begin
 result:=10;
 Q:=MOpenSql('SELECT MAX(QBH_NUMNOEUD)+1 FROM QBPARBREDETAIL '+
             'WHERE QBH_CODESESSION="'+codeSession+'"',
             'BPFctArbre (BPIncrementenumNoeudDetail).',true);
 if not Q.eof
  then
   begin
    if Q.fields[0].asString<>''
     then result:=Q.fields[0].asInteger;
   end;
 ferme(Q);
end;


//*****************************************************************************
//Vide Arbre

//vide complétement l'arbre de la session donnée
procedure VideArbre(const codeSession:hString;Update:boolean);
begin
//Booléen pour récupération des anciennes session
if Update=false then MExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"',
                                 'BPFctArbre (VideArbre).');
MExecuteSql('DELETE FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"',
             'BPFctArbre (VideArbre).');
end;

//vide l'arbre de la session donnée à partir du niveau donné
procedure VideArbreNiveau(const codeSession:hString;Niveau:integer);
begin
 MExecuteSql('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NIVEAU>="'+STRFPOINT(niveau)+'"',
             'BPFctArbre (VideArbreNiveau).');
end;

//*****************************************************************************
//Mise à jour

//-----------------------
//Mise à jour des quantités prévues et retenues
//-----------------------

//Mise à jour des qtes prevues et retenues à partir du dernier niveau de l'arbre de prevision
procedure MiseAJourQtePrevueQteRetenue(const codeSession:hString);
var i,Niveau:integer;
begin
 Niveau:=ChercheNivMaxSession(codeSession);

 for i:=Niveau-1 downto 1 do
  begin
   MExecuteSql('UPDATE QBPARBRE SET QBR_PREVU='+
               '(SELECT SUM(QTEPREVU) FROM QUVBPVUEFILSARBRE WHERE '+
               'S="'+codeSession+'" AND NIV="'+STRFPOINT(i+1)+
               '" AND NOEUDPERE=QBR_NUMNOEUD) '+
               'WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NIVEAU="'+STRFPOINT(i)+'"',
               'BPFctArbre (MiseAJourQtePrevueQteRetenue).');     
   MExecuteSql('UPDATE QBPARBRE SET QBR_QTERETENUE='+
               '(SELECT SUM(QTERET) FROM QUVBPVUEFILSARBRE WHERE '+
               'S="'+codeSession+'" AND NIV="'+STRFPOINT(i+1)+
               '" AND NOEUDPERE=QBR_NUMNOEUD) '+
               'WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NIVEAU="'+STRFPOINT(i)+'"',
               'BPFctArbre (MiseAJourQtePrevueQteRetenue).');
  end;
end;

//-----------------------
//Mise à jour noeud pere
//-----------------------

function CodeSqlMAJNoeudPere(Chp1,Chp2:hString):hString;
begin
 result:=' ( '+Chp1+'='+Chp2+' OR ('+Chp1+' is null and '+Chp2+' is null) '+
         ' OR ('+Chp1+' is null and '+Chp2+'="") '+
         ' OR ('+Chp1+' ="" and '+Chp2+' is null) )';
end;

//Mise à jour des noeuds pères suivant les noeuds definis
//sachant qu'au départ tous les noeuds sont à 0
procedure MiseAJourNoeudPere(const codeSession:hString);
var i,j,Niveau:integer;
    codeWhere:hString;
begin
 if SessionInitPrev(codeSession)
  then Niveau:=ChercheNivMaxSession(codeSession)
  else
   begin
    Niveau:=ChercheNivMax(codeSession);
    if SessionDelai(codeSession)
     then Niveau:=Niveau+1;
   end;

  for i:=2 to Niveau do
  begin
    if not MoveCurProgressForm(TraduireMemoire('Validation...')) then
    begin
       MoveCurProgressForm(TraduireMemoire('Annulation en cours...'));
       break;
    end;

    codeWhere:=CodeSqlMAJNoeudPere('VALEURAXE','QBR_VALAXENIV'+IntToStr(i-1));
    for j:=2 to i-1 do
    codeWhere:=codeWhere+' AND '+CodeSqlMAJNoeudPere(' VALNIV'+IntToStr(j-1),'QBR_VALAXENIV'+IntToStr(j-1));

    MExecuteSql('UPDATE QBPARBRE SET QBR_NUMNOEUDPERE='+
                '(SELECT NOEUD FROM QUVBPVUEFILSARBRE WHERE '+
                'S="'+codeSession+'" AND NIV="'+STRFPOINT(i-1)+
                '" AND '+codeWhere+' AND DEV=QBR_DEVISE) '+
                'WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NIVEAU="'+STRFPOINT(i)+'"',
                'BPFctArbre (MiseAJourNoeudPere).');
     {codeWhere:='A.QBR_VALEURAXE=R.QBR_VALAXENIV'+IntToStr(i-1);
     for j:=2 to i-1 do
     codeWhere:=codeWhere+' AND A.QBR_VALAXENIV'+IntToStr(j-1)+'=R.QBR_VALAXENIV'+IntToStr(j-1);
      MExecuteSql('UPDATE QBPARBRE R SET R.QBR_NUMNOEUDPERE='+
                 '(SELECT A.QBR_NUMNOEUD FROM QBPARBRE A WHERE '+
                 'A.QBR_CODESESSION="'+codeSession+'" AND A.QBR_NIVEAU="'+STRFPOINT(i-1)+
                 '" AND '+codeWhere+') '+
                 'WHERE R.QBR_CODESESSION="'+codeSession+'" AND R.QBR_NIVEAU="'+STRFPOINT(i)+'"',
                 'BPFctArbre (MiseAJourNoeudPere).');}
  end;
end;

//-----------------------
//Mise à jour %
//-----------------------

//requete type pour mettre à jour des % dans l'arbre (qbparbre)
procedure RequeteMajPrctArbre(const codeSession,codePrct,codeVal,codeValVue,niveau:hString);
var codeMaj:hString;
begin
 codeMAJ:=''+codePrct+'='+
          'IIF((SELECT SUM('+codeValVue+') FROM QUVBPVUEFILSARBRE WHERE '+
          'S="'+codeSession+
          '" AND NIV="'+niveau+
          '" AND NOEUDPERE=QBR_NUMNOEUDPERE)=0,'+
          '(SELECT 100.0/COUNT(*) FROM QUVBPVUEFILSARBRE WHERE '+
          'S="'+codeSession+
          '" AND NIV="'+niveau+
          '" AND NOEUDPERE=QBR_NUMNOEUDPERE),'+
          codeVal+'/'+
          '(SELECT SUM('+codeValVue+') FROM QUVBPVUEFILSARBRE WHERE '+
          'S="'+codeSession+
          '" AND NIV="'+niveau+
          '" AND NOEUDPERE=QBR_NUMNOEUDPERE)*100) ';
 MExecuteSql('UPDATE QBPARBRE SET '+codeMAJ+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NIVEAU="'+niveau+'"',
             'BPFctArbre (RequeteMajPrctArbre).');

{ codeMAJ:='R.'+codePrct+'='+
          'IIF((SELECT SUM(A.'+codeVal+') FROM QBPARBRE A WHERE '+
          'A.QBR_CODESESSION="'+codeSession+
          '" AND A.QBR_NIVEAU="'+niveau+
          '" AND A.QBR_NUMNOEUDPERE=R.QBR_NUMNOEUDPERE)=0,0,R.'+codeVal+'/'+
          '(SELECT SUM(A.'+codeVal+') FROM QBPARBRE A WHERE '+
          'A.QBR_CODESESSION="'+codeSession+
          '" AND A.QBR_NIVEAU="'+niveau+
          '" AND A.QBR_NUMNOEUDPERE=R.QBR_NUMNOEUDPERE)*100) ';
 MExecuteSql('UPDATE QBPARBRE R SET '+codeMAJ+
             'WHERE R.QBR_CODESESSION="'+codeSession+
             '" AND R.QBR_NIVEAU="'+niveau+'"',
             'BPFctArbre (RequeteMajPrctArbre).');  }
end;

//Mise à jour des % dans arbre
//si OkMajPrctRef alors maj des % reference
//si OkMajPrctObj alors maj des % objectif prevision
procedure MiseAJourPrctArbre(const codeSession:hString;okMajPrctRef,OkMajPrctObj:boolean);
var i,j,Niveau:integer;
    codePrct,codeVal,codeValVue:hString;
begin
  Niveau:=ChercheNivMaxSession(codeSession);
  if SessionDelai(codeSession) then Niveau:=Niveau+1;
  for i:=1 to Niveau do
  begin
    if not MoveCurProgressForm(TraduireMemoire('Validation...')) then
    begin
      MoveCurProgressForm(TraduireMemoire('Annulation en cours...'));
      break;
    end;

    for j:=1 to 7 do
    begin
      if okMajPrctRef then
      begin
        codePrct:='QBR_REFPRCT'+IntToStr(j);
        codeVal:='QBR_REF'+IntToStr(j);
        codeValVue:='REF'+IntToStr(j);
        if j=7 then
        begin
          //pour la quantité
          codePrct:='QBR_QTEREFPRCT';
          codeVal:='QBR_QTEREF';
          codeValVue:='QTEREF';
        end;
        RequeteMajPrctArbre(codeSession,codePrct,codeVal,codeValVue,STRFPOINT(i));
      end;
      if okMajPrctObj then
      begin
        codePrct:='QBR_OPPRCT'+IntToStr(j);
        codeVal:='QBR_OP'+IntToStr(j);
        codeValVue:='OP'+IntToStr(j);
        if j=7 then
        begin
          //pour la quantité
          codePrct:='QBR_QTECPRCT';
          codeVal:='QBR_QTEC';
          codeValVue:='QTEC';
        end;
        RequeteMajPrctArbre(codeSession,codePrct,codeVal,codeValVue,STRFPOINT(i));
      end
    end
  end
end;

//Mise à jour des % dans arbre pour le noeud pere donné
//si OkMajPrctRef alors maj des % reference
//si OkMajPrctObj alors maj des % objectif prevision
procedure MiseAJourPrctArbreNoeudPere(const codeSession,NoeudPere:hString;
          okMajPrctRef,OkMajPrctObj:boolean;
          TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
var codeMaj,codePrct,codeVal:hString;
    j:integer;
    TabVal:array [1..7] of double;
begin
 TabVal[1]:=TotalCA1;
 TabVal[2]:=TotalCA2;
 TabVal[3]:=TotalCA3;
 TabVal[4]:=TotalCA4;
 TabVal[5]:=TotalCA5;
 TabVal[6]:=TotalCA6;
 TabVal[7]:=TotalQte;

 for j:=1 to 7 do
  begin
   if okMajPrctRef
    then
     begin
      codePrct:='QBR_REFPRCT'+IntToStr(j);
      codeVal:='QBR_REF'+IntToStr(j);
      if j=7
       then
        begin
         //pour la quantité
         codePrct:='QBR_QTEREFPRCT';
         codeVal:='QBR_QTEREF';
        end;
       codeMAJ:=codePrct+'='+
                'IIF("'+STRFPOINT(TabVal[j])+'"="0",0,'+codeVal+'/"'+STRFPOINT(TabVal[j])+'"*100) ';
       MExecuteSql('UPDATE QBPARBRE SET '+codeMAJ+
                   'WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NUMNOEUDPERE="'+NoeudPere+'"',
                   'BPFctArbre (MiseAJourPrctArbreNoeudPere).');
     end;
   if okMajPrctObj
    then
     begin
      codePrct:='QBR_OPPRCT'+IntToStr(j);
      codeVal:='QBR_OP'+IntToStr(j);
      if j=7
       then
        begin
         //pour la quantité
         codePrct:='QBR_QTECPRCT';
         codeVal:='QBR_QTEC';
        end;
       codeMAJ:=codePrct+'='+
                'IIF("'+STRFPOINT(TabVal[j])+'"="0",0,'+codeVal+'/"'+STRFPOINT(TabVal[j])+'"*100) ';
       MExecuteSql('UPDATE QBPARBRE SET '+codeMAJ+
                   'WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NUMNOEUDPERE="'+NoeudPere+'"',
                   'BPFctArbre (MiseAJourPrctArbreNoeudPere).');
     end;
   end;
end;

//-----------------------
//Mise à jour niveau
//-----------------------


//mise a jour des niveaux de l'arbre
procedure MiseAjourNiveauSuivant(MAJTable:boolean;numnoeud,codeSession:hString;
          niveau,nivMaxSession,RI:integer;newCAVal1,newQteVal,
          newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,TotalRet,CARetenu:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
var Q:TQuery;
    val,Qteval:double;
    CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double;
    numnoeudQ:hString;
    TOBNoeud:TOB;
    i,k:integer;
    NoeudFils:hString;
    Ref1,Ref2,Ref3,Ref4,Ref5,Ref6,RefQte:double;
    OPPRCT1,OPPRCT2,OPPRCT3,OPPRCT4,OPPRCT5,OPPRCT6,OPPRCTQte:double;
    OP1,OP2,OP3,OP4,OP5,OP6,OPQte:double;
    OPQteArrondi,QteRetenuArrondi,QteRealise:double;
    QteRetenu:double;
    errOPQte,errQteRetenu,rapport,somme:double;
    ListeQte,ListeQteRes:array[0..19] of double;
    valBloque,Modif0:boolean;
    SomFrereCAVal1,SomFrereQteval,SomFrereCAVal2,SomFrereCAVal3:double;
    SomFrereCAVal4,SomFrereCAVal5,SomFrereCAVal6:double;     
    BQteval,BCAVal1,BCAVal2,BCAVal3,BCAVal4,BCAVal5,BCAVal6:double;
begin
 Modif0:=false;
 //si niveau>Niveau max de la session
 //ou niveau>10 on sort de la recursivité
 if (niveau>10) or (niveau>nivMaxSession)
  then exit;

 //si dernier niveau et init taille
 //mise à jour de la grille de taille
 if (OkSessionInitPrev) and (niveau=nivMaxSession) and (OkInitNivTaille)
  then
   begin
    TOBNoeud:=TOB.Create('ARBRE',nil,-1);
    Q:=MOPenSql('SELECT * FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUD="'+numnoeud+'"','',true);
    TOBNoeud.LoadDetailDB( 'QBPARBRE', '', 'BPFctArbre (MiseAjourNiveauSuivant).', Q, False );
    Ferme( Q );
    for i:=0 to TOBNoeud.detail.count-1 do
     begin
      NoeudFils:=TOBNoeud.detail[i].getvalue('QBR_NUMNOEUD');

      rapport:=0;
      somme:=0;
      for k:=1 to 20 do
       begin
        ListeQte[k-1]:= valeurR(TOBNoeud.detail[i].getvalue('QBR_QTET'+IntToStr(k)));
        somme:=somme+ListeQte[k-1];
       end;

      if somme<>0
       then rapport:=(TOBNoeud.detail[i].getvalue('QBR_QTERETENUE'))/somme;

      BPCalculQteListProrata(0,ListeQte,Rapport,ListeQteRes);

      for k:=1 to 20 do
       begin
        TOBNoeud.detail[i].putvalue('QBR_QTET'+IntToStr(k),floatToStr(ListeQteRes[k-1]));
       end;

      TOBNoeud.UpdateDB(false);
     end;
    TobNoeud.free;
   end;


 TOBNoeud:=TOB.Create('ARBRE',nil,-1);

 { EVI / Flag pour bug pourcentage d'évolution = 0 }
 Q:=MOPenSql('SELECT QBR_VALMODIF FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+numnoeud+'"','BPFctArbre (MiseAjourNiveauSuivant).',true);
 if not Q.eof then
 begin
   if Q.fields[0].asString='X' then Modif0:=true;
 end;
 Ferme( Q );

 Q:=MOPenSql('SELECT * FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+numnoeud+'"','BPFctArbre (MiseAjourNiveauSuivant).',true);
 TOBNoeud.LoadDetailDB( 'QBPARBRE', '', '', Q, False );
 Ferme( Q );
 errOPQte:=0;
 errQteRetenu:=0;
 for i:=0 to TOBNoeud.detail.count-1 do
  begin
   { EVI / Flag pour bug pourcentage d'évolution = 0 }
   if Modif0=true then TOBNoeud.detail[i].putvalue('QBR_VALMODIF','X');



   NoeudFils:=TOBNoeud.detail[i].getvalue('QBR_NUMNOEUD');
   Ref1:=TOBNoeud.detail[i].getvalue('QBR_REF1');
   Ref2:=TOBNoeud.detail[i].getvalue('QBR_REF2');
   Ref3:=TOBNoeud.detail[i].getvalue('QBR_REF3');
   Ref4:=TOBNoeud.detail[i].getvalue('QBR_REF4');
   Ref5:=TOBNoeud.detail[i].getvalue('QBR_REF5');
   Ref6:=TOBNoeud.detail[i].getvalue('QBR_REF6');
   RefQte:=TOBNoeud.detail[i].getvalue('QBR_QTEREF');
   
   OP1:=TOBNoeud.detail[i].getvalue('QBR_OP1');
   OP2:=TOBNoeud.detail[i].getvalue('QBR_OP2');
   OP3:=TOBNoeud.detail[i].getvalue('QBR_OP3');
   OP4:=TOBNoeud.detail[i].getvalue('QBR_OP4');
   OP5:=TOBNoeud.detail[i].getvalue('QBR_OP5');
   OP6:=TOBNoeud.detail[i].getvalue('QBR_OP6');
   OPQte:=TOBNoeud.detail[i].getvalue('QBR_QTEC');

   OPPRCT1:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT1');
   OPPRCT2:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT2');
   OPPRCT3:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT3');
   OPPRCT4:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT4');
   OPPRCT5:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT5');
   OPPRCT6:=TOBNoeud.detail[i].getvalue('QBR_OPPRCT6');
   OPPRCTQte:=TOBNoeud.detail[i].getvalue('QBR_QTECPRCT');

   QteRealise:=TOBNoeud.detail[i].getvalue('QBR_REALISE');
                                     
   ValBloque:=(TOBNoeud.detail[i].getvalue('QBR_VALBLOQUETMP')='X');


   //si il existe un frere bloqué
   if ExisteSql('SELECT QBR_VALBLOQUETMP FROM QBPARBRE '+
           ' WHERE QBR_CODESESSION="'+codeSession+
           '" AND QBR_NUMNOEUDPERE="'+numnoeud+
           '" AND QBR_VALBLOQUETMP="X" ')
    then
     begin
      //somme des frères non bloqués
      SomFrereCAVal1:=1;
      SomFrereQteval:=1;
      SomFrereCAVal2:=1;
      SomFrereCAVal3:=1;
      SomFrereCAVal4:=1;
      SomFrereCAVal5:=1;
      SomFrereCAVal6:=1;
      Q:=MOpenSql('SELECT SUM(QBR_REF1),SUM(QBR_QTEREF),'+
                  'SUM(QBR_REF2),SUM(QBR_REF3),SUM(QBR_REF4),'+
                  'SUM(QBR_REF5),SUM(QBR_REF6) FROM QBPARBRE '+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUDPERE="'+numnoeud+
                  '" AND QBR_VALBLOQUETMP="-" ',
                  'BPFctArbre (MiseAJourNiveauSuivantValBloquee)',true);
      if not Q.eof
       then
        begin
         SomFrereCAVal1:=Q.fields[0].asFloat;
         SomFrereQteval:=Q.fields[1].asFloat;
         SomFrereCAVal2:=Q.fields[2].asFloat;
         SomFrereCAVal3:=Q.fields[3].asFloat;
         SomFrereCAVal4:=Q.fields[4].asFloat;
         SomFrereCAVal5:=Q.fields[5].asFloat;
         SomFrereCAVal6:=Q.fields[6].asFloat;
        end;
      ferme(Q);

      //somme des frères bloqués
      BCAVal1:=0;
      BQteval:=0;
      BCAVal2:=0;
      BCAVal3:=0;
      BCAVal4:=0;
      BCAVal5:=0;
      BCAVal6:=0;
      Q:=MOpenSql('SELECT SUM(QBR_OP1),SUM(QBR_QTEC),'+
                  'SUM(QBR_OP2),SUM(QBR_OP3),SUM(QBR_OP4),'+
                  'SUM(QBR_OP5),SUM(QBR_OP6) FROM QBPARBRE '+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUDPERE="'+numnoeud+
                  '" AND QBR_VALBLOQUETMP="X" ',
                  'BPFctArbre (MiseAJourNiveauSuivantValBloquee)',true);
      if not Q.eof
       then
        begin
         BCAVal1:=Q.fields[0].asFloat;
         BQteval:=Q.fields[1].asFloat;
         BCAVal2:=Q.fields[2].asFloat;
         BCAVal3:=Q.fields[3].asFloat;
         BCAVal4:=Q.fields[4].asFloat;
         BCAVal5:=Q.fields[5].asFloat;
         BCAVal6:=Q.fields[6].asFloat;
        end;
      ferme(Q);

      DiffCa1:=newCAVal1-SomFrereCAVal1-BCAVal1;
      DiffQte:=newQteVal-SomFrereQteVal-BQteVal;
      DiffCa2:=newCAVal2-SomFrereCAVal2-BCAVal2;
      DiffCa3:=newCAVal2-SomFrereCAVal2-BCAVal3;
      DiffCa4:=newCAVal2-SomFrereCAVal2-BCAVal4;
      DiffCa5:=newCAVal2-SomFrereCAVal2-BCAVal5;
      DiffCa6:=newCAVal2-SomFrereCAVal2-BCAVal6;

      //si la valeur est non bloquée
      if not valBloque
       then
        begin
         //CA1
         if SomFrereCAVal1=0
          then OP1:=0
          else OP1:=Ref1+(Ref1/SomFrereCAVal1)*DiffCa1;
         if newCAVal1=0
          then OPPRCT1:=0
          else OPPRCT1:=(OP1/newCAVal1)*100;
         //CA2
         if SomFrereCAVal2=0
          then OP2:=0
          else OP2:=Ref2+(Ref2/SomFrereCAVal2)*DiffCa2;
         if newCAVal2=0
          then OPPRCT2:=0
          else OPPRCT2:=(OP2/newCAVal2)*100;
         //CA3
         if SomFrereCAVal3=0
          then OP3:=0
          else OP3:=Ref3+(Ref3/SomFrereCAVal3)*DiffCa3;
         if newCAVal3=0
          then OPPRCT3:=0
          else OPPRCT3:=(OP3/newCAVal3)*100;
         //CA4
         if SomFrereCAVal4=0
          then OP4:=0
          else OP4:=Ref4+(Ref4/SomFrereCAVal4)*DiffCa4;
         if newCAVal4=0
          then OPPRCT4:=0
          else OPPRCT4:=(OP4/newCAVal4)*100;
         //CA5
         if SomFrereCAVal5=0
          then OP5:=0
          else OP5:=Ref5+(Ref5/SomFrereCAVal5)*DiffCa5;
         if newCAVal5=0
          then OPPRCT5:=0
          else OPPRCT5:=(OP5/newCAVal5)*100;
         //CA6
         if SomFrereCAVal6=0
          then OP6:=0
          else OP6:=Ref6+(Ref6/SomFrereCAVal6)*DiffCa6;
         if newCAVal6=0
          then OPPRCT6:=0
          else OPPRCT6:=(OP6/newCAVal6)*100;
         //Qte
         if SomFrereQteval=0
          then OPQte:=0
          else OPQte:=RefQte+(RefQte/SomFrereQteval)*DiffQte;
         if newQteVal=0
          then OPPRCTQte:=0
          else OPPRCTQte:=(OPQte/newQteVal)*100;
        end
       else
        begin
         //CA1
         if newCAVal1=0
          then OPPRCT1:=0
          else OPPRCT1:=(OP1/newCAVal1)*100;
         //CA2
         if newCAVal2=0
          then OPPRCT2:=0
          else OPPRCT2:=(OP2/newCAVal2)*100;
         //CA3
         if newCAVal3=0
          then OPPRCT3:=0
          else OPPRCT3:=(OP3/newCAVal3)*100;
         //CA4
         if newCAVal4=0
          then OPPRCT4:=0
          else OPPRCT4:=(OP4/newCAVal4)*100;
         //CA5
         if newCAVal5=0
          then OPPRCT5:=0
          else OPPRCT5:=(OP5/newCAVal5)*100;
         //CA6
         if newCAVal6=0
          then OPPRCT6:=0
          else OPPRCT6:=(OP6/newCAVal6)*100;
         //Qte
         if newQteVal=0
          then OPPRCTQte:=0
          else OPPRCTQte:=(OPQte/newQteVal)*100;
        end;

      TOBNoeud.detail[i].putvalue('QBR_OPPRCT1',OPPRCT1);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCT2',OPPRCT2);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCT3',OPPRCT3);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCT4',OPPRCT4);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCT5',OPPRCT5);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCT6',OPPRCT6);
      TOBNoeud.detail[i].putvalue('QBR_OPPRCTQTE',OPPRCTQTE);

     end
    else
     begin
      //si la valeur est non bloquée
      OP1:=(OPPRCT1*newCAVal1)/100;
      OP2:=(OPPRCT2*newCAVal2)/100;
      OP3:=(OPPRCT3*newCAVal3)/100;
      OP4:=(OPPRCT4*newCAVal4)/100;
      OP5:=(OPPRCT5*newCAVal5)/100;
      OP6:=(OPPRCT6*newCAVal6)/100;
      OPQte:=(OPPRCTQte*newQteVal)/100;
     end;

   if OkModifCoeff
    then QteRetenu:=(valeurR(newCoeff)*QteRealise)
    else QteRetenu:=(OPPRCTQte*TotalRet)/100;

   //pour les qtés calcul des arrondis
   OPQteArrondi:=round(OPQte+errOPQte);
   errOPQte:=errOPQte+(OPQte-OPQteArrondi);

   QteRetenuArrondi:=round(QteRetenu+errQteRetenu);
   errQteRetenu:=errQteRetenu+(QteRetenu-QteRetenuArrondi);

   //si la valeur est non bloquée
   if not valBloque then
   begin
     //MAj des valeurs
     if ContextBP = 1 then
     begin
       case RI of
        1 : TOBNoeud.detail[i].putvalue('QBR_OP1',OP1);
        2 : TOBNoeud.detail[i].putvalue('QBR_QTEC',OPQteArrondi);
        3 : TOBNoeud.detail[i].putvalue('QBR_OP2',OP2);
        4 : TOBNoeud.detail[i].putvalue('QBR_OP3',OP3);
        5 : TOBNoeud.detail[i].putvalue('QBR_OP4',OP4);
        6 : TOBNoeud.detail[i].putvalue('QBR_OP5',OP5);
        7 : TOBNoeud.detail[i].putvalue('QBR_OP6',OP6);
       end;
     end
     else
     begin
       TOBNoeud.detail[i].putvalue('QBR_OP1',OP1);
       TOBNoeud.detail[i].putvalue('QBR_QTEC',OPQteArrondi);
       TOBNoeud.detail[i].putvalue('QBR_OP2',OP2);
       TOBNoeud.detail[i].putvalue('QBR_OP3',OP3);
       TOBNoeud.detail[i].putvalue('QBR_OP4',OP4);
       TOBNoeud.detail[i].putvalue('QBR_OP5',OP5);
       TOBNoeud.detail[i].putvalue('QBR_OP6',OP6);
       TOBNoeud.detail[i].putvalue('QBR_QTERETENUE',QteRetenuArrondi);
     end;

      //si modif coeff
      if OkModifCoeff
       then
        begin
         TOBNoeud.detail[i].putvalue('QBR_COEFFRETENU',valeurR(newCoeff));
         TOBNoeud.detail[i].putvalue('QBR_CARETENU',TOBNoeud.detail[i].getvalue('QBR_REALISECA')*valeurR(newCoeff));
         TOBNoeud.detail[i].putvalue('QBR_PREVU',QteRetenuArrondi);
         TOBNoeud.detail[i].putvalue('QBR_PREVUCA',TOBNoeud.detail[i].getvalue('QBR_REALISECA')*valeurR(newCoeff));
        end;
     end;

  TOBNoeud.UpdateDB(false);

  MiseAjourNiveauSuivant(MAJTable,NoeudFils,codeSession,niveau+1,nivMaxSession,RI,
                          OP1,OPQteArrondi,OP2,OP3,OP4,OP5,OP6,
                          OPQteArrondi-RefQte,OP1-Ref1,OP2-Ref2,OP3-Ref3,OP4-Ref4,OP5-Ref5,OP6-Ref6,
                          QteRetenuArrondi,CARetenu,
                          newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
  end;

 TobNoeud.free;
end;

procedure MiseAjourNiveauPrecedentValNonBloquee(MAJTable:boolean;
          niveau,nivMaxSession:integer;
          codeSession,numNoeudPere,NumNoeud:hString;
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
          TotCAVal1,TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,totQteVal,
          TotalPrevu,TotalRet,CARetenu:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
var cont:boolean;
    Q:TQuery;
    numnoeudQ,code:hString;
    TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalP,TotalR,TotalRCA:double;
begin
 code:='';
 if OkModifCoeff
  then code:='" ,QBR_CARETENU="'+strFPoint(CARetenu);
 //mise à jour du niveau courant
 MExecuteSql('UPDATE QBPARBRE SET '+
             'QBR_OP1="'+strFPoint(TotCAVal1)+
             '" ,QBR_QTEC="'+strFPoint(totQteVal)+
             '" ,QBR_OP2="'+strFPoint(TotCAVal2)+
             '" ,QBR_OP3="'+strFPoint(TotCAVal3)+
             '" ,QBR_OP4="'+strFPoint(TotCAVal4)+
             '" ,QBR_OP5="'+strFPoint(TotCAVal5)+
             '" ,QBR_OP6="'+strFPoint(TotCAVal6)+
             '" ,QBR_PREVU="'+strFPoint(TotalPrevu)+
             '" ,QBR_QTERETENUE="'+strFPoint(TotalRet)+
             code+
             '" WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+numnoeudpere+'"',
             'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');

 cont:=false;

 //cherche le noeud pere du niveau courant
 Q:=MOpenSql('SELECT QBR_NUMNOEUDPERE FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+numnoeudpere+'"',
             'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).',true);
 if not Q.eof
  then
   begin
    numnoeudQ:=Q.fields[0].asString;
    cont:=true;
   end;
 ferme(Q);

 if cont
  then
   begin
    ArbreTotalNiv(codeSession,numnoeudQ,
                  TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
                  TotalP,TotalR,TotalRCA);

    //mise à jour du niveau courant "niveau"
    //attention
    if TotalCA1<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT1=((QBR_OP1/'+strFPoint(TotalCA1)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalQte<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_QTECPRCT=((QBR_QTEC/'+strFPoint(TotalQte)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalCA2<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT2=((QBR_OP2/'+strFPoint(TotalCA2)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalCA3<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT3=((QBR_OP3/'+strFPoint(TotalCA3)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalCA4<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT4=((QBR_OP4/'+strFPoint(TotalCA4)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalCA5<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT5=((QBR_OP5/'+strFPoint(TotalCA5)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');
    if TotalCA6<>0
     then MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_OPPRCT6=((QBR_OP6/'+strFPoint(TotalCA6)+')*100) '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUDPERE="'+numnoeudQ+'"',
            'BPFctArbre (MiseAjourNiveauPrecedentValNonBloquee).');

    //lancement récursif avec le nouveau noeud père trouvé
    //avec niveau -1
    //et avec nouvelle valeur pour le noeud pere du niveau précédent : totalniv
    MiseAjourNiveauPrecedent(MAJTable,codeSession,numnoeudpere,numnoeudQ,
                             niveau-1,nivMaxSession,totalCA1,TotalQte,
                             TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
                             DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
                             TotalP,TotalR,TotalRCA,
                             newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
   end;

end;

//
procedure MiseAjourNiveauPrecedentValBloquee(const codeSession,numNoeudPere,NumNoeud:hString;
          MAJTable:boolean;
          Niveau,nivMaxSession:integer;
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
var TotFrereQte,TotFrereCA,TotFrereCA2,TotFrereCA3,TotFrereCA4,TotFrereCA5,TotFrereCA6:double;
    codeSql,noeud:hString;
    Q:TQuery;
    RI:integer;
    newCAVal1,newQteVal,newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6:double;
    CAVal1,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,QteRetenue,CARetenu:double;

    function Coma(St: String): String;
    begin
      if St='' then result:='' else result:=',';
    end;

begin
 //Totaux des freres
 TotFrereCA:=1;
 TotFrereQte:=1;
 TotFrereCA2:=1;
 TotFrereCA3:=1;
 TotFrereCA4:=1;
 TotFrereCA5:=1;
 TotFrereCA6:=1;
 Q:=MOpenSQL('SELECT SUM(QBR_OP1),SUM(QBR_QTEC),SUM(QBR_OP2),'+
             'SUM(QBR_OP3),SUM(QBR_OP4),SUM(QBR_OP5),'+
             'SUM(QBR_OP6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+numNoeudPere+
             '" AND QBR_NUMNOEUD<>"'+NumNoeud+
             '" AND QBR_VALBLOQUETMP="-"',
             'BPFctArbre (MiseAjourNiveauPrecedentValBloquee).',true);
 if not Q.eof
  then
   begin
    TotFrereCA:=Q.fields[0].asFloat;
    TotFrereQte:=Q.fields[1].asFloat;
    TotFrereCA2:=Q.fields[2].asFloat;
    TotFrereCA3:=Q.fields[3].asFloat;
    TotFrereCA4:=Q.fields[4].asFloat;
    TotFrereCA5:=Q.fields[5].asFloat;
    TotFrereCA6:=Q.fields[6].asFloat;
   end;
 ferme(Q);

 //maj des freres non bloqués
 if ContextBP in [0,1,2] then
 begin
   RI:=DonneValeurAffiche(codeSession);
   case RI of
     1 : codeSql:='QBR_OP1=(QBR_OP1-(QBR_OP1/'+strFPoint(TotFrereCA)+')*'+strFPoint(DiffCA1)+')  ';
     2 : codeSql:='QBR_QTEC=(QBR_QTEC-(QBR_QTEC/'+strFPoint(TotFrereQte)+')*'+strFPoint(DiffQte)+') ';
     3 : codeSql:='QBR_OP2=(QBR_OP2-(QBR_OP2/'+strFPoint(TotFrereCA2)+')*'+strFPoint(DiffCA2)+') ';
     4 : codeSql:='QBR_OP3=(QBR_OP3-(QBR_OP3/'+strFPoint(TotFrereCA3)+')*'+strFPoint(DiffCA3)+') ';
     5 : codeSql:='QBR_OP4=(QBR_OP4-(QBR_OP4/'+strFPoint(TotFrereCA4)+')*'+strFPoint(DiffCA4)+') ';
     6 : codeSql:='QBR_OP5=(QBR_OP5-(QBR_OP5/'+strFPoint(TotFrereCA5)+')*'+strFPoint(DiffCA5)+') ';
     7 : codeSql:='QBR_OP6=(QBR_OP6-(QBR_OP6/'+strFPoint(TotFrereCA6)+')*'+strFPoint(DiffCA6)+') ';
   end;
 end
 else
 begin
   RI:=0;
   if TotFrereCA<>0 then codeSql:='QBR_OP1=(QBR_OP1-(QBR_OP1/'+strFPoint(TotFrereCA)+')*'+strFPoint(DiffCA1)+')  ';
   if TotFrereQte<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_QTEC=(QBR_QTEC-(QBR_QTEC/'+strFPoint(TotFrereQte)+')*'+strFPoint(DiffQte)+') ';
   if TotFrereCA2<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_OP2=(QBR_OP2-(QBR_OP2/'+strFPoint(TotFrereCA2)+')*'+strFPoint(DiffCA2)+') ';
   if TotFrereCA3<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_OP3=(QBR_OP3-(QBR_OP3/'+strFPoint(TotFrereCA3)+')*'+strFPoint(DiffCA3)+') ';
   if TotFrereCA4<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_OP4=(QBR_OP4-(QBR_OP4/'+strFPoint(TotFrereCA4)+')*'+strFPoint(DiffCA4)+') ';
   if TotFrereCA5<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_OP5=(QBR_OP5-(QBR_OP5/'+strFPoint(TotFrereCA5)+')*'+strFPoint(DiffCA5)+') ';
   if TotFrereCA6<>0 then codeSql:=codeSql+Coma(CodeSQL)+'QBR_OP6=(QBR_OP6-(QBR_OP6/'+strFPoint(TotFrereCA6)+')*'+strFPoint(DiffCA6)+') ';

 end;

 if OkModifCoeff
  then codeSql:=codeSql+',QBR_PREVU=(QBR_REALISE*'+NewCoeff+
                '),QBR_QTERETENUE=(QBR_REALISE*'+NewCoeff+') ';

 if codeSql<>'' then MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
                                 'WHERE QBR_CODESESSION="'+codeSession+
                                 '" AND QBR_NUMNOEUDPERE="'+numNoeudPere+
                                 '" AND QBR_NUMNOEUD<>"'+NumNoeud+
                                 '" AND QBR_VALBLOQUETMP="-"',
                                 'BPFctArbre (MiseAjourNiveauPrecedentValBloquee).');
 //maj des fils des freres
 Q:=MOPenSql('SELECT QBR_NUMNOEUD FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUDPERE="'+numNoeudPere+
                '" AND QBR_NUMNOEUD<>"'+NumNoeud+
                '" AND QBR_VALBLOQUETMP="-"',
                'BPFctArbre (MiseAjourNiveauPrecedentValBloquee).',true);
  while not Q.eof do
   begin
    noeud:=Q.fields[0].asString;
    ArbreValeurNoeud(codeSession,noeud,
                     newCAVal1,newQteVal,newCAVal2,newCAVal3,
                     newCAVal4,newCAVal5,newCAVal6,QteRetenue,CARetenu);
    ArbreValeurRefNoeud(codeSession,noeud,
                        CAVal1,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6);
    MiseAjourNiveauSuivant(MAJTable,Noeud,codeSession,niveau,nivMaxSession,
                       RI,newCAVal1,newQteVal,
                       newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,
                       abs(DiffQte),abs(DiffCA1),abs(DiffCA2),abs(DiffCA3),
                       abs(DiffCA4),abs(DiffCA5),abs(DiffCA6),QteRetenue,CARetenu,
                       newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
    Q.next
   end;
  ferme(Q);

end;


//procedure recursive
//qui met à jour les niveaux précèdents
procedure MiseAjourNiveauPrecedent(MAJTable:boolean;codeSession,numNoeud,numnoeudpere:hString;
          niveau,nivMaxSession:integer;TotCAVal1,totQteVal,
          TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
          TotalPrevu,TotalRet,CARetenu:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
begin
 //arrêt de la récursivité
 //au niveau 0
 if niveau<=0
  then exit;

 if ExisteSql('SELECT QBR_VALBLOQUETMP FROM QBPARBRE '+
              ' WHERE QBR_CODESESSION="'+codeSession+
              '" AND QBR_NUMNOEUD="'+numnoeudpere+
              '" AND QBR_VALBLOQUETMP="X" ')
  then MiseAjourNiveauPrecedentValBloquee(codeSession,numNoeudPere,NumNoeud,
                                          MAJTable,Niveau,nivMaxSession,
                                          DiffQte,DiffCA1,DiffCA2,DiffCA3,
                                          DiffCA4,DiffCA5,DiffCA6,
                                          newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille)
  else MiseAjourNiveauPrecedentValNonBloquee(MAJTable,Niveau,nivMaxSession,
                                          codeSession,numNoeudPere,NumNoeud,
                                          DiffQte,DiffCA1,DiffCA2,DiffCA3,
                                          DiffCA4,DiffCA5,DiffCA6,
                                          TotCAVal1,TotCAVal2,TotCAVal3,
                                          TotCAVal4,TotCAVal5,TotCAVal6,totQteVal,
                                          TotalPrevu,TotalRet,CARetenu,
                                          newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
end;

//procedure qui met à jour l'arbre suivant les modifs d'un noeud donné
procedure MiseAjourNiv(MAJTable:boolean;const codeSession,Noeud,NoeudPere,Niveau:hString;
          nivMaxSession:integer;
          TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
          DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6:double;
          newCoeff:hString;OkModifCoeff,OkSessionInitPrev,OkInitNivTaille:boolean);
var newVal,newQteVal,totVal,totQteVal:double;
    newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,CARetenu:double;
    TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,TotalPrevu,TotalRet,QteRetenue,TotalCARetenu:double;
    RI:integer;
begin
  if MAJTable
   then
    begin
     MiseAJourPrctArbreNoeudPere(codeSession,noeudPere,false,true,
                                    TotalQte,TotalCA1,TotalCA2,TotalCA3,
                                    TotalCA4,TotalCA5,TotalCA6);
     if OkSessionInitPrev
      then
       begin                               
        ArbreTotalNiv(codeSession,noeud,
                      totQteVal,totVal,TotCAVal2,TotCAVal3,
                      TotCAVal4,TotCAVal5,TotCAVal6,TotalPrevu,TotalRet,TotalCARetenu);
        MiseAJourPrctArbreNoeudPere(codeSession,noeud,false,true,
                                    totQteVal,totVal,TotCAVal2,TotCAVal3,
                                    TotCAVal4,TotCAVal5,TotCAVal6);
       end;
    end;

  ArbreValeurNoeud(codeSession,noeud,
                   newVal,newQteVal,newCAVal2,newCAVal3,
                   newCAVal4,newCAVal5,newCAVal6,QteRetenue,CARetenu);

  ArbreTotalNiv(codeSession,noeudpere,
                totQteVal,totVal,TotCAVal2,TotCAVal3,
                TotCAVal4,TotCAVal5,TotCAVal6,TotalPrevu,TotalRet,TotalCARetenu);
  
  RI:=DonneValeurAffiche(codeSession);
  if BPOkOrli
   then
    //-----------------> ORLI
    begin
     if DiffQte<>0
      then RI:=2;
     if DiffCA1<>0
      then RI:=1;
     if DiffCA2<>0
      then RI:=3;
     if DiffCA3<>0
      then RI:=4;
     if DiffCA4<>0
      then RI:=5;
     if DiffCA5<>0
      then RI:=6;
     if DiffCA6<>0
      then RI:=7;
    end;
    //ORLI <-----------------

  MiseAjourNiveauSuivant(MAJTable,Noeud,codeSession,VALEURI(niveau),nivMaxSession,
                         RI,newVal,newQteVal,
                         newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,
                         DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
                         QteRetenue,CARetenu,
                         newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
  MiseAjourNiveauPrecedent(MAJTable,codeSession,Noeud,NoeudPere,VALEURI(niveau),nivMaxSession,
                           totVal,totQteVal,
                           TotCAVal2,TotCAVal3,TotCAVal4,TotCAVal5,TotCAVal6,
                           DiffQte,DiffCA1,DiffCA2,DiffCA3,DiffCA4,DiffCA5,DiffCA6,
                           TotalPrevu,TotalRet,TotalCARetenu,
                           newCoeff,OkModifCoeff,OkSessionInitPrev,OkInitNivTaille);
end;

//*****************************************************************************
//Valeur bloquée


procedure MiseAJourValeurBloque(const codeSession:hString;NivMax:integer);
var Nb:integer;
begin
 MExecuteSql('UPDATE QBPARBRE SET QBR_VALBLOQUETMP=QBR_VALBLOQUE '+
             'WHERE QBR_CODESESSION="'+codeSession+'"',
             'BPFctArbre (MiseAJourValeurBloque).');
 MExecuteSql('UPDATE QBPARBRE SET QBR_VALBLOQUETMP1="-" '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_VALBLOQUETMP1<>"-" ',
             'BPFctArbre (MiseAJourValeurBloque).');
 Nb:=1;
 while Nb<>0 do
  Nb:=MExecuteSql('UPDATE QBPARBRE SET QBR_VALBLOQUETMP="X",QBR_VALBLOQUETMP1="X" '+
             'WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_VALBLOQUETMP="-" '+
             'AND '+
             'EXISTS (SELECT NOEUD FROM QUVBPVUEFILSARBRE WHERE S="'+codeSession+
             '" AND NOEUDPERE=QBR_NUMNOEUD) AND '+
             'NOT EXISTS (SELECT NOEUD FROM QUVBPVUEVALBLOQARBRE WHERE S="'+codeSession+
             '" AND NOEUDPERE=QBR_NUMNOEUD AND VALBLOQ="-" )',
             'BPFctArbre (MiseAJourValeurBloque).')+
       MExecuteSql('UPDATE QBPARBRE SET QBR_VALBLOQUETMP="X",QBR_VALBLOQUETMP1="X" '+
             'WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_VALBLOQUETMP="-" '+
             'AND '+
             'EXISTS (SELECT NOEUD FROM QUVBPVUEVALBLOQARBRE WHERE S="'+codeSession+
             '" AND NOEUD=QBR_NUMNOEUDPERE and VALBLOQ="X") AND '+
             'NOT EXISTS (SELECT NOEUD FROM QUVBPVUEVALBLOQARBRE WHERE S="'+codeSession+
             '" AND NOEUDPERE=QBR_NUMNOEUDPERE AND NOEUD<>QBR_NUMNOEUD AND VALBLOQ="-" )',
             'BPFctArbre (MiseAJourValeurBloque).');
end;

function NoeudCourantValBloquee(const codeSession,numNoeud,numNoeudPere:hString):boolean;
var OkFils,OkPere:boolean;
begin
 OkFils:=false;
 OkPere:=false;
 //bloque si le noeud est bloque par heritage
 if ExisteSql('SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+numNoeud+'" AND QBR_VALBLOQUETMP1="X"')
  then OkFils:=true;
 //si le pere est bloque
 if ExisteSql('SELECT QBR_VALBLOQUETMP FROM QBPARBRE '+
              ' WHERE QBR_CODESESSION="'+codeSession+
              '" AND QBR_NUMNOEUD="'+numNoeudPere+
              '" AND QBR_VALBLOQUETMP="X" ')
  then
   begin
    //si il n'existe pas au moins un frere non bloque
    if not ExisteSql('SELECT QBR_NUMNOEUD FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUDPERE="'+numNoeudPere+
                '" AND QBR_NUMNOEUD<>"'+numNoeud+
                '" AND QBR_VALBLOQUETMP="-"')
     then OkPere:=true;
   end;

 result:=OkFils or OkPere;
end;

//*****************************************************************************
//Total niveau

//retourne les totaux des réferences du niveau
procedure ArbreTotalNivRefPere(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
var Q:TQuery;
begin
 TotalQte:=1;
 TotalCA1:=1;
 TotalCA2:=1;
 TotalCA3:=1;
 TotalCA4:=1;
 TotalCA5:=1;  
 TotalCA6:=1;
 //total niveau
{ Q:=MOpenSql('SELECT SUM(QBR_QTEREF),SUM(QBR_REF1),SUM(QBR_REF2),'+
             'SUM(QBR_REF3),SUM(QBR_REF4),SUM(QBR_REF5),'+
             'SUM(QBR_REF6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNivRefPere).',true);
} Q:=MOpenSql('SELECT SUM(QBR_QTEC),SUM(QBR_OP1),SUM(QBR_OP2),'+
             'SUM(QBR_OP3),SUM(QBR_OP4),SUM(QBR_OP5),'+
             'SUM(QBR_OP6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNivRefPere).',true);
 if not Q.eof
  then
   begin
    TotalQte:=Q.fields[0].asFloat;
    TotalCA1:=Q.fields[1].asFloat;
    TotalCA2:=Q.fields[2].asFloat;
    TotalCA3:=Q.fields[3].asFloat;
    TotalCA4:=Q.fields[4].asFloat;
    TotalCA5:=Q.fields[5].asFloat;
    TotalCA6:=Q.fields[6].asFloat;
   end;
 ferme(Q);
end;

//retourne les totaux des réferences du niveau
procedure ArbreTotalNivRef(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
var Q:TQuery;
begin
 TotalQte:=1;
 TotalCA1:=1;
 TotalCA2:=1;
 TotalCA3:=1;
 TotalCA4:=1;
 TotalCA5:=1;  
 TotalCA6:=1;
 //total niveau
{ Q:=MOpenSql('SELECT SUM(QBR_QTEREF),SUM(QBR_REF1),SUM(QBR_REF2),'+
             'SUM(QBR_REF3),SUM(QBR_REF4),SUM(QBR_REF5),'+
             'SUM(QBR_REF6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNivRef).',true);      }
 
 Q:=MOpenSql('SELECT SUM(QBR_QTEC),SUM(QBR_OP1),SUM(QBR_OP2),'+
             'SUM(QBR_OP3),SUM(QBR_OP4),SUM(QBR_OP5),'+
             'SUM(QBR_OP6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNivRef).',true);
 if not Q.eof
  then
   begin
    TotalQte:=Q.fields[0].asFloat;
    TotalCA1:=Q.fields[1].asFloat;
    TotalCA2:=Q.fields[2].asFloat;
    TotalCA3:=Q.fields[3].asFloat;
    TotalCA4:=Q.fields[4].asFloat;
    TotalCA5:=Q.fields[5].asFloat;
    TotalCA6:=Q.fields[6].asFloat;
   end;
 ferme(Q);
end;

//retourne les totaux des valeurs courantes du niveau
procedure ArbreTotalNiv(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
          TotalPrevu,TotalRet,CARetenu:double);
var Q:TQuery;
begin
 TotalQte:=1;
 TotalCA1:=1;
 TotalCA2:=1;
 TotalCA3:=1;
 TotalCA4:=1;
 TotalCA5:=1;  
 TotalCA6:=1;
 TotalPrevu:=1;
 TotalRet:=1;
 CARetenu:=1;
 //total niveau
 Q:=MOpenSql('SELECT SUM(QBR_QTEC),SUM(QBR_OP1),SUM(QBR_OP2),'+
             'SUM(QBR_OP3),SUM(QBR_OP4),SUM(QBR_OP5),'+
             'SUM(QBR_OP6),SUM(QBR_PREVU),SUM(QBR_QTERETENUE),'+
             'SUM(QBR_CARETENU) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNiv).',true);
 if not Q.eof
  then
   begin
    TotalQte:=Q.fields[0].asFloat;
    TotalCA1:=Q.fields[1].asFloat;
    TotalCA2:=Q.fields[2].asFloat;
    TotalCA3:=Q.fields[3].asFloat;
    TotalCA4:=Q.fields[4].asFloat;
    TotalCA5:=Q.fields[5].asFloat;
    TotalCA6:=Q.fields[6].asFloat;    
    TotalPrevu:=Q.fields[7].asFloat;
    TotalRet:=Q.fields[8].asFloat;
   end;
 ferme(Q);
end;

//retourne les totaux des valeurs courantes du niveau
procedure ArbreTotalNivHisto(const codeSession,noeudPere:hString;
          var TotalQte,TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6:double);
var Q:TQuery;
begin
 TotalQte:=1;
 TotalCA1:=1;
 TotalCA2:=1;
 TotalCA3:=1;
 TotalCA4:=1;
 TotalCA5:=1;
 TotalCA6:=1;
 //total niveau
 Q:=MOpenSql('SELECT SUM(QBR_QTEREF),SUM(QBR_REF1),SUM(QBR_REF2),'+
             'SUM(QBR_REF3),SUM(QBR_REF4),SUM(QBR_REF5),'+
             'SUM(QBR_REF6) FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"',
             'BPFctArbre (ArbreTotalNivHisto).',true);
 if not Q.eof
  then
   begin
    TotalQte:=Q.fields[0].asFloat;
    TotalCA1:=Q.fields[1].asFloat;
    TotalCA2:=Q.fields[2].asFloat;
    TotalCA3:=Q.fields[3].asFloat;
    TotalCA4:=Q.fields[4].asFloat;
    TotalCA5:=Q.fields[5].asFloat;
    TotalCA6:=Q.fields[6].asFloat;
   end;
 ferme(Q);
end;

//*****************************************************************************
//Total arbre

//cherche la valeur totale
//pour le type de champ :
//CA1,Qte,CA2,CA3,CA4,CA5
function ChercheTotalArbre(const codeSession,typeC:hString):double;
var Q:TQuery;
    total:double;
    codeSql:hString;
begin
 total:=0;
 if typeC='CA1'
  then codeSql:='QBR_OP1';
 if typeC='QTE'
  then codeSql:='QBR_QTEC';
 if typeC='CA2'
  then codeSql:='QBR_OP2';
 if typeC='CA3'
  then codeSql:='QBR_OP3';
 if typeC='CA4'
  then codeSql:='QBR_OP4';
 if typeC='CA5'
  then codeSql:='QBR_OP5';
 if typeC='CA6'
  then codeSql:='QBR_OP6';

 Q:=MOPenSql('SELECT SUM('+codeSql+') FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="0"',
             'BPFctArbre (ChercheTotalArbre)',true);
 if not Q.eof
  then total:=Q.fields[0].asfloat;
 ferme(Q);
 result:=total;
end;


//*****************************************************************************
//Total noeud

            
//retourne les valeurs reference du noeud
procedure ArbreValeurNoeud(const codeSession,noeud:hString;
          var newVal,newQteVal,newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,QteRetenue,CARetenu:double);
var Q:TQuery;
begin
 newVal:=0;
 newQteVal:=0;
 newCAVal2:=0;
 newCAVal3:=0;
 newCAVal4:=0;
 newCAVal5:=0;
 newCAVal6:=0;   
 QteRetenue:=0;
 CARetenu:=0;
 //new valeur courante
 Q:=MOpenSql('SELECT QBR_OP1,QBR_QTEC,'+
            'QBR_OP2,QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6,QBR_QTERETENUE,QBR_CARETENU '+
            'FROM QBPARBRE '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUD="'+Noeud+'"',
            'BPFctArbre (ArbreValeurNoeud)',true);
 if not Q.eof
  then
   begin
    newVal:=Q.fields[0].asFloat;
    newQteVal:=Q.fields[1].asFloat;
    newCAVal2:=Q.fields[2].asFloat;
    newCAVal3:=Q.fields[3].asFloat;
    newCAVal4:=Q.fields[4].asFloat;
    newCAVal5:=Q.fields[5].asFloat;
    newCAVal6:=Q.fields[6].asFloat;
    QteRetenue:=Q.fields[7].asFloat;       
    CARetenu:=Q.fields[8].asFloat;
   end;
 ferme(Q);
end;
                         
//retourne les valeurs courantes du noeud
procedure ArbreValeurRefNoeud(const codeSession,noeud:hString;
          var Val,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double);
var Q:TQuery;
begin
 Val:=0;
 QteVal:=0;
 CAVal2:=0;
 CAVal3:=0;
 CAVal4:=0;
 CAVal5:=0;
 CAVal6:=0;
 //new valeur courante
 Q:=MOpenSql('SELECT QBR_REF1,QBR_QTEREF,QBR_REF2,'+
             'QBR_REF3,QBR_REF4,QBR_REF5,'+
             'QBR_REF6 '+
             'FROM QBPARBRE '+
             ' WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+Noeud+'"',
             'BPFctArbre (ArbreValeurRefNoeud).',true);
 if not Q.eof
  then
   begin
    Val:=Q.fields[0].asFloat;
    QteVal:=Q.fields[1].asFloat;
    CAVal2:=Q.fields[2].asFloat;
    CAVal3:=Q.fields[3].asFloat;
    CAVal4:=Q.fields[4].asFloat;
    CAVal5:=Q.fields[5].asFloat;
    CAVal6:=Q.fields[6].asFloat;
   end;
 ferme(Q);
end;

//retourne les valeurs courantes du noeud
procedure ArbreValeurRefNoeudTotal(const codeSession:hString;
          var Val,QteVal,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6:double);
var Q:TQuery;
begin
 Val:=0;
 QteVal:=0;
 CAVal2:=0;
 CAVal3:=0;
 CAVal4:=0;
 CAVal5:=0;
 CAVal6:=0;
 //new valeur courante
 Q:=MOpenSql('SELECT QBR_REF1,QBR_QTEREF,QBR_REF2,'+
             'QBR_REF3,QBR_REF4,QBR_REF5,'+
             'QBR_REF6 '+
             'FROM QBPARBRE '+
             ' WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="0"',
             'BPFctArbre (ArbreValeurRefNoeudTotal).',true);
 if not Q.eof
  then
   begin
    Val:=Q.fields[0].asFloat;
    QteVal:=Q.fields[1].asFloat;
    CAVal2:=Q.fields[2].asFloat;
    CAVal3:=Q.fields[3].asFloat;
    CAVal4:=Q.fields[4].asFloat;
    CAVal5:=Q.fields[5].asFloat;
    CAVal6:=Q.fields[6].asFloat;
   end;
 ferme(Q);
end;

//*****************************************************************************
//Controle arbre

procedure ControleNoeudANoeudPere(const codeSession:hString);
var niv,i:integer;
    Q:TQuery;
    Ok:boolean;
begin
 niv:=ChercheNivMaxSession(codeSession);
 Ok:=true;
 for i:=1 to niv-1 do
  begin
   Q:=MOpenSql('select qbr_numnoeud from qbparbre '+
               'where qbr_codesession="'+codeSession+
               '" and qbr_niveau="'+intToStr(i+1)+
               '" and qbr_numnoeudpere not in '+
               '(select qbr_numnoeud from qbparbre '+
               'where qbr_codesession="'+codeSession+
               '" and qbr_niveau="'+intToStr(i)+'")',
               'BPFctArbre (ControleNoeudANoeudPere)',true);
   while not Q.eof do
    begin
     Ok:=false;
     MExecuteSql('delete from qbparbre '+
                'where qbr_codesession="'+codeSession+
                '" and qbr_numnoeud="'+Q.fields[0].asString+'"',
                'BPFctArbre (ControleNoeudANoeudPere).');
     Q.next;
    end;
   ferme(Q);
  end;
 if ok=false
  then  PGIINFO(Format_(
         traduireMemoire('Suppression dans l''arbre (session : %s) de noeud sans noeud père. Appeller l''assistance.'),
               [codeSession]));
end;

//*****************************************************************************
//Evolution

procedure EvolutionSurUneValeurAxe(const retour,codeSession,codeaxe,valeuraxe:hString;Onglet:integer);
var Q:TQuery;
    noeud,noeudpere,niveau:hString;
    evolprct,saival,evolval,codeSql:hString;
    RI,nivMaxSession:integer;
    ChpHisto,ChpNew:hString;
    okMo,OkSessionInitPrev,OkInitNivTaille:boolean;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu:double;
   { EvolQteVal,EvolQtePrct,SaisieQteVal:hString;
    EvolCAVal2,EvolCAPrct2,SaisieCAVal2:hString;
    EvolCAVal3,EvolCAPrct3,SaisieCAVal3:hString;
    EvolCAVal4,EvolCAPrct4,SaisieCAVal4:hString;
    EvolCAVal5,EvolCAPrct5,SaisieCAVal5:hString; }
begin
 //Vérification valeurs bloquées
 if ExisteSQL('SELECT QBR_VALBLOQUE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_VALBLOQUE="X"') then
 begin
   PGIInfo('Traitement impossible : il existe des valeurs bloquées.');
   exit;
 end;

 //CA1
 evolprct:=trouveargument(retour,'PRCT','0');
 saival:=trouveargument(retour,'SAISIE','0');
 evolval:=trouveargument(retour,'VAL','0');
{ //Qte
 EvolQteVal:=trouveargument(retour,'EVOLQTEVAL','0');
 EvolQtePrct:=trouveargument(retour,'EVOLQTEPRCT','0');
 SaisieQteVal:=trouveargument(retour,'SAISIEQTEVAL','0');
 //CA2
 EvolCAVal2:=trouveargument(retour,'EVOLCAVAL2','0');
 EvolCAPrct2:=trouveargument(retour,'EVOLCAPRCT2','0');
 SaisieCAVal2:=trouveargument(retour,'SAISIECAVAL2','0');
 //CA3
 EvolCAVal3:=trouveargument(retour,'EVOLCAVAL3','0');
 EvolCAPrct3:=trouveargument(retour,'EVOLCAPRCT3','0');
 SaisieCAVal3:=trouveargument(retour,'SAISIECAVAL3','0');
 //CA4
 EvolCAVal4:=trouveargument(retour,'EVOLCAVAL4','0');
 EvolCAPrct4:=trouveargument(retour,'EVOLCAPRCT4','0');
 SaisieCAVal4:=trouveargument(retour,'SAISIECAVAL4','0');
 //CA5
 EvolCAVal5:=trouveargument(retour,'EVOLCAVAL5','0');
 EvolCAPrct5:=trouveargument(retour,'EVOLCAPRCT5','0');
 SaisieCAVal5:=trouveargument(retour,'SAISIECAVAL5','0'); }

 //cherche le code de la requête
 codeSql:=''; RI := 0;
 Case ContextBP of
   0,1,2 : RI:=DonneValeurAffiche(codeSession);
   3 : RI := Onglet;
 end;

 case RI of
   1 : begin
         ChpHisto:='QBR_REF1';
         ChpNew:=' QBR_OP1';
       end;
   2 : begin
         ChpHisto:='QBR_QTEREF';
         ChpNew:=' QBR_QTEC';
       end;
   3 : begin
         ChpHisto:='QBR_REF2';
         ChpNew:=' QBR_OP2';
       end;
   4 : begin
         ChpHisto:='QBR_REF3';
         ChpNew:=' QBR_OP3';
       end;
   5 : begin
         ChpHisto:='QBR_REF4';
         ChpNew:=' QBR_OP4';
       end;
   6 : begin
         ChpHisto:='QBR_REF5';
         ChpNew:=' QBR_OP5';
       end;
   7 : begin
         ChpHisto:='QBR_REF6';
         ChpNew:=' QBR_OP6';
       end;
 end;

 okMo:=false;
 if saival<>''
  then codeSql:=ChpNew+'="'+STRFPOINT(VALEUR(saival))+'" '
  else
   begin
    if (evolprct<>'') and (evolval<>'')
     then
      begin
       codeSql:=ChpNew+'='+ChpHisto+'+(('+ChpHisto+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
       okMo:=true;
      end;
    if (evolprct<>'') and (evolval='')
     then codeSql:=ChpNew+'='+ChpHisto+'+(('+ChpHisto+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
    if (okMo=false) and (evolval<>'')
     then codeSql:=ChpNew+'='+ChpHisto+'+ "'+STRFPOINT(VALEUR(evolval))+'" ';
   end;
 {
 //qte
 if SaisieQteVal<>''
  then
   begin
    if codesql=''
     then codeSql:=' QBI_NEWQTEVAL="'+SaisieQteVal+'" '
     else codeSql:=codesql+' ,QBI_NEWQTEVAL="'+SaisieQteVal+'" ';
   end
  else
   begin
    if (EvolQtePrct<>'') and (EvolQteVal<>'')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWQTEVAL=QBI_QTEVAL+((QBI_QTEVAL*"'+EvolQtePrct+'")/100) + "'+EvolQteVal+'" '
        else codeSql:=codesql+' ,QBI_NEWQTEVAL=QBI_QTEVAL+((QBI_QTEVAL*"'+EvolQtePrct+'")/100) + "'+EvolQteVal+'" ';
      end;
    if (EvolQtePrct<>'') and (EvolQteVal='')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWQTEVAL=QBI_QTEVAL+((QBI_QTEVAL*"'+EvolQtePrct+'")/100) '
        else codeSql:=codesql+' ,QBI_NEWQTEVAL=QBI_QTEVAL+((QBI_QTEVAL*"'+EvolQtePrct+'")/100) ';
      end;
   end;
 //CA2
 if SaisieCAVal2<>''
  then
   begin
    if codesql=''
     then codeSql:=' QBI_NEWCAVAL2="'+SaisieCAVal2+'" '
     else codeSql:=codesql+' ,QBI_NEWCAVAL2="'+SaisieCAVal2+'" ';
   end
  else
   begin
    if (EvolCAPrct2<>'') and (EvolCAVal2<>'')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL2=QBI_CAVAL2+((QBI_CAVAL2*"'+EvolCAPrct2+'")/100) + "'+EvolCAVal2+'" '
        else codeSql:=codesql+' ,QBI_NEWCAVAL2=QBI_CAVAL2+((QBI_CAVAL2*"'+EvolCAPrct2+'")/100) + "'+EvolCAVal2+'" ';
      end;
    if (EvolCAPrct2<>'') and (EvolCAVal2='')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL2=QBI_CAVAL2+((QBI_CAVAL2*"'+EvolCAPrct2+'")/100) '
        else codeSql:=codesql+' ,QBI_NEWCAVAL2=QBI_CAVAL2+((QBI_CAVAL2*"'+EvolCAPrct2+'")/100) ';
      end;
   end;
 //CA3
 if SaisieCAVal3<>''
  then
   begin
    if codesql=''
     then codeSql:=' QBI_NEWCAVAL3="'+SaisieCAVal3+'" '
     else codeSql:=codesql+' ,QBI_NEWCAVAL3="'+SaisieCAVal3+'" ';
   end
  else
   begin
    if (EvolCAPrct3<>'') and (EvolCAVal3<>'')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL3=QBI_CAVAL3+((QBI_CAVAL3*"'+EvolCAPrct3+'")/100) + "'+EvolCAVal3+'" '
        else codeSql:=codesql+' ,QBI_NEWCAVAL3=QBI_CAVAL3+((QBI_CAVAL3*"'+EvolCAPrct3+'")/100) + "'+EvolCAVal3+'" ';
      end;
    if (EvolCAPrct3<>'') and (EvolCAVal3='')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL3=QBI_CAVAL3+((QBI_CAVAL3*"'+EvolCAPrct3+'")/100) '
        else codeSql:=codesql+' ,QBI_NEWCAVAL3=QBI_CAVAL3+((QBI_CAVAL3*"'+EvolCAPrct3+'")/100) ';
      end;
   end;
 //CA4
 if SaisieCAVal4<>''
  then
   begin
    if codesql=''
     then codeSql:=' QBI_NEWCAVAL4="'+SaisieCAVal4+'" '
     else codeSql:=codesql+' ,QBI_NEWCAVAL4="'+SaisieCAVal4+'" ';
   end
  else
   begin
    if (EvolCAPrct4<>'') and (EvolCAVal4<>'')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL4=QBI_CAVAL4+((QBI_CAVAL4*"'+EvolCAPrct4+'")/100) + "'+EvolCAVal4+'" '
        else codeSql:=codesql+' ,QBI_NEWCAVAL4=QBI_CAVAL4+((QBI_CAVAL4*"'+EvolCAPrct4+'")/100) + "'+EvolCAVal4+'" ';
      end;
    if (EvolCAPrct4<>'') and (EvolCAVal4='')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL4=QBI_CAVAL4+((QBI_CAVAL4*"'+EvolCAPrct4+'")/100) '
        else codeSql:=codesql+' ,QBI_NEWCAVAL4=QBI_CAVAL4+((QBI_CAVAL4*"'+EvolCAPrct4+'")/100) ';
      end;
   end;
 //CA5
 if SaisieCAVal5<>''
  then
   begin
    if codesql=''
     then codeSql:=' QBI_NEWCAVAL5="'+SaisieCAVal5+'" '
     else codeSql:=codesql+' ,QBI_NEWCAVAL5="'+SaisieCAVal5+'" ';
   end
  else
   begin
    if (EvolCAPrct5<>'') and (EvolCAVal5<>'')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL5=QBI_CAVAL5+((QBI_CAVAL5*"'+EvolCAPrct5+'")/100) + "'+EvolCAVal5+'" '
        else codeSql:=codesql+' ,QBI_NEWCAVAL5=QBI_CAVAL5+((QBI_CAVAL5*"'+EvolCAPrct5+'")/100) + "'+EvolCAVal5+'" ';
      end;
    if (EvolCAPrct5<>'') and (EvolCAVal5='')
     then
      begin
       if codesql=''
        then codeSql:=' QBI_NEWCAVAL5=QBI_CAVAL5+((QBI_CAVAL5*"'+EvolCAPrct5+'")/100) '
        else codeSql:=codesql+' ,QBI_NEWCAVAL5=QBI_CAVAL5+((QBI_CAVAL5*"'+EvolCAPrct5+'")/100) ';
      end;
   end;
  }
  
 //maj de qbparbre
 if (codeSql<>'')
  then
   begin
    //met à jour la table
    MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
               ' WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_CODEAXE="'+codeaxe+
               '" AND QBR_VALEURAXE="'+valeuraxe+'" AND QBR_VALBLOQUE="-"',
               'BPFctArbre (EvolutionSurUneValeurAxe).');
   end;

 nivMaxSession:=ChercheNivMaxSession(codeSession);
 OkSessionInitPrev:=SessionInitPrev(codeSession);   
 OkInitNivTaille:=(SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession));

 Q:=MOPenSql('SELECT QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_NIVEAU '+
             'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_CODEAXE="'+codeaxe+
             '" AND QBR_VALEURAXE="'+valeuraxe+'" AND QBR_VALBLOQUE="-"',
             'BPFctArbre (EvolutionSurUneValeurAxe)',true);
 while not Q.eof do
  begin
   noeud:=Q.fields[0].asstring;
   noeudpere:=Q.fields[1].asstring;
   niveau:=Q.fields[2].asstring;
   ArbreTotalNiv(codeSession,noeudPere,
                 TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                 TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
   MiseAjourNiv(true,codeSession,Noeud,NoeudPere,Niveau,nivMaxSession,
                TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                TotalNCA5,TotalNCA6,0,0,0,0,0,0,0,'0',false,OkSessionInitPrev,OkInitNivTaille);
   Q.next;
  end;
 ferme(Q);

 MAJPrctVariation(codeSession);
end;

//*****************************************************************************
//MetArbreEnLigne

//procedure
procedure MetArbreEnLigneNiveau(Niveau:integer;const types,session,ChpsqlR,codeFiltre:hString;
          TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalQte:double;
          TabCodeAxe:array of hString;QTBP:TQTob;okCoeff,okPrev:boolean;nbTabAffiche:integer);
var Q:TQuery;
    codesql,chpSql,joinSql:hString;
    i: integer;
//    ii:integer;
    TabValAxe,TabValAxeLib:array [1..12] of variant;
    noeud,valeur:hString;
    CalculDelaiOk:boolean;
    TabD:array of double;
    TabChpVal:array of variant;
    NbChp,nbChpSql:integer;
    okSessionInitPrev: boolean;
    AxeSalarie:integer;
    TabCodNewSal,TabLibNewSal: array of string;
begin
  CalculDelaiOk:= SessionDelai(session);
  okSessionInitPrev:= SessionInitPrev(session);
  codesql:=ChpSqlR;

  for i:=1 to Niveau-1 do codesql:=codesql+',QBR_VALAXENIV'+IntToStr(i);

  ChercheCodeSqlChpAxeLib(Niveau,TabCodeAxe, chpSql,joinSql);
  if (CalculDelaiOk) and (not okSessionInitPrev)
  then chpSql:=chpSql+' ,QBR_VALEURAXE '
  else chpSql:=chpSql+' ,QBR_LIBVALAXE ';

  nbChpSql:=nbTabAffiche*4;
  if OkCoeff then nbChpSql:=31;
  if OkPrev then nbChpSql:=3;
  SetLength(TabD,nbChpSql);

  { EVI / Libelle pour les Nouvelles Embauches en Paie }
  AxeSalarie := 0;
  if ContextBP = 3 then
  begin
    for i:=low(TabCodeAxe) to high(TabCodeAxe) do
    begin
      if TabCodeAxe[i]='011' then
      begin
        AxeSalarie := i;
        break;
      end;
    end;
  end;
  if AxeSalarie <> 0 then
  begin
    Q:=OpenSql('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PTK"',true);
    SetLength(TabCodNewSal,Q.RecordCount);
    SetLength(TabLibNewSal,Q.RecordCount);
    i:=0;
    while not Q.eof do
    begin
      TabCodNewSal[i]:=Q.fields[0].asString;
      TabLibNewSal[i]:=Q.fields[1].asString;
      i:=i+1;
      Q.next;
    end;
    ferme(Q);
  end;

  Q:=MOPenSql('SELECT QBR_NUMNOEUD,QBR_VALEURAXE,'+codesql+chpsql+' FROM QBPARBRE '+
              joinsql+' WHERE QBR_CODESESSION="'+session+
              '" AND QBR_NIVEAU="'+IntToStr(Niveau)+'" '+codeFiltre,
              'BPFctArbre (MetArbreEnLigneNiveau).',true);

  InitMoveProgressForm(nil, TraduireMemoire('Contrôle session') , TraduireMemoire('Veuillez patienter'),Q.RecordCount , True, True);
  while not Q.eof do
  begin
    if not MoveCurProgressForm(TraduireMemoire('Génération...')) then
    begin
       MoveCurProgressForm(TraduireMemoire('Annulation en cours...'));
       break;
    end;
    noeud:=Q.fields[0].asString;
    valeur:=Q.fields[1].asString;
    for i:=0 to nbChpSql-1 do TabD[i]:=Q.fields[2+i].asFloat;

    for i:=1 to Niveau-1 do TabValAxe[i]:=Q.fields[nbChpSql+i+1].asString;
    if (CalculDelaiOk) and (not okSessionInitPrev) then TabValAxe[Niveau]:=StrToDateTime(valeur)
    else TabValAxe[Niveau]:=valeur;

    for i:=1 to Niveau-1 do TabValAxeLib[i]:=Q.fields[nbChpSql+Niveau+i].asString;
    if (CalculDelaiOk) and (not okSessionInitPrev) then TabValAxeLib[Niveau]:=valeur
    else TabValAxeLib[Niveau]:=valeur;

    { EVI / Libelle pour les Nouvelles Embauches en Paie }
    If AxeSalarie <> 0 then
    begin
      for i:=low(TabCodNewSal) to high(TabCodNewSal) do
      begin
        if TabValAxe[AxeSalarie]=TabCodNewSal[i] then
        begin
          TabValAxeLib[AxeSalarie]:=TabLibNewSal[i];
          break;
        end
      end;
    end;

    {if Niveau=1
    then TabValAxeLib[ii]:=Q.fields[nbChpSql+Niveau+ii].asString
    else TabValAxeLib[ii+1]:=Q.fields[nbChpSql+Niveau+ii+1].asString;}

    NbChp:=(2*Niveau);
    SetLength(TabChpVal,NbChp+1);
    for i:=1 to Niveau do TabChpVal[i-1]:=TabValAxe[i];
    for i:=Niveau to Niveau*2 do TabChpVal[i]:=TabValAxeLib[i-Niveau+1];
    TabChpVal[Niveau*2]:=noeud;

    QTBP.addValeur(TabChpVal,TabD);

    Q.next;
  end;
  ferme(Q);
  if not MoveCurProgressForm(TraduireMemoire('Génération...')) then QTBP.latob.clearDetail;
  FiniMoveProgressForm
end;



//procedure
procedure MetArbreEnLigne(NivMaxProfil:integer;const types,session:hString;
          total:double;
          TabCodeAxe:array of hString;QTBP:TQTob);
var Q:TQuery;
    codesql,chpSql,joinSql:hString;
    i:integer;
    TabValAxe,TabValAxeLib:array [1..12] of hString;
    noeud,valeur:hString;
    CalculDelaiOk:boolean;
    BPInitialise:hString;
    CAVal1,CAVal2,CAVal3,CAVal4,CAVal5,CAVal6,QteVal:double;
    NewCAVal1,NewCAVal2,NewCAVal3,NewCAVal4,NewCAVal5,NewCAVal6,NewQteVal:double;
    OkModif:boolean;
    RI:integer;
begin
 CalculDelaiOk:=SessionCalculParDelai(session,BPInitialise);
 codesql:='QBR_NUMNOEUD,QBR_VALEURAXE,'+
          'QBR_REF1,QBR_OP1,QBR_REF2,QBR_OP2,'+
          'QBR_REFL3,QBR_OP3,QBR_REF4,QBR_OP4,'+
          'QBR_REF5,QBR_OPL5,QBR_REF6,QBR_OP6,QBR_QTEREF,QBR_QTEC ';

 if (types='VALIDATION') and CalculDelaiOk
  then NivMaxProfil:=NivMaxProfil+1;

 for i:=1 to NivMaxProfil-1 do
  codesql:=codesql+',QBR_VALAXENIV'+IntToStr(i);

 ChercheCodeSqlChpAxeLib(NivMaxProfil,TabCodeAxe,
                         chpSql,joinSql);

 RI:=DonneValeurAffiche(session);

 Q:=MOPenSql('SELECT '+codesql+chpsql+' FROM QBPARBRE '+
             joinsql+' WHERE QBR_CODESESSION="'+session+
             '" AND QBR_NIVEAU="'+IntToStr(NivMaxProfil)+'"',
             'BPFctArbre (MetArbreEnLigne).',true);
 while not Q.eof do
  begin
   noeud:=Q.fields[0].asString;
   valeur:=Q.fields[1].asString;
   CAVAl1:=Q.fields[2].asFloat;
   NewCAVAl1:=Q.fields[3].asFloat;
   CAVAl2:=Q.fields[4].asFloat;
   NewCAVAl2:=Q.fields[5].asFloat;
   CAVAl3:=Q.fields[6].asFloat;
   NewCAVAl3:=Q.fields[7].asFloat;
   CAVAl4:=Q.fields[8].asFloat;
   NewCAVAl4:=Q.fields[9].asFloat;
   CAVAl5:=Q.fields[10].asFloat;
   NewCAVAl5:=Q.fields[11].asFloat;  
   CAVAl6:=Q.fields[12].asFloat;
   NewCAVAl6:=Q.fields[13].asFloat;
   QteVAl:=Q.fields[14].asFloat;
   NewQteVAl:=Q.fields[15].asFloat;

   for i:=1 to NivMaxProfil-1 do
    TabValAxe[i]:=Q.fields[15+i].asString;
   TabValAxe[NivMaxProfil]:=valeur;
   for i:=NivMaxProfil to NivMaxProfil+NivMaxProfil-2 do
    TabValAxeLib[i-NivMaxProfil+1]:=Q.fields[15+i].asString;

  OkModif:=false;
  case RI of
    1 : begin
        if (NewCAVal1<>CAVal1)
         then OkModif:=true;
        end;
    2 : begin
        if (NewQteVal<>QteVal)
         then OkModif:=true;
        end;
    3 : begin
        if (NewCAVal2<>CAVal2)
         then OkModif:=true;
        end;
    4 : begin
        if (NewCAVal3<>CAVal3)
         then OkModif:=true;
        end;
    5 : begin
        if (NewCAVal4<>CAVal4)
         then OkModif:=true;
        end;
    6 : begin
        if (NewCAVal5<>CAVal5)
         then OkModif:=true;
        end; 
    7 : begin
        if (NewCAVal6<>CAVal6)
         then OkModif:=true;
        end;
    end;

   if OkModif
    then QTBP.addValeur([TabValAxe[1],TabValAxe[2],TabValAxe[3],
                              TabValAxe[4],TabValAxe[5],TabValAxe[6],
                              TabValAxe[7],TabValAxe[8],TabValAxe[9],
                              TabValAxe[10],noeud],
                              [CAVal1,NewCAVal1,CAVal2,NewCAVal2,
                               CAVal3,NewCAVal3,CAVal4,NewCAVal4,
                               CAVal5,NewCAVal5,CAVal6,NewCAVal6,QteVal,NewQteVal]);
   Q.next;
  end;
 ferme(Q);
end;

//*****************************************************************************
//Noeud pere

//retourne le numéro de noeud pere pour un noeud courant donné
function ChercheNoeudPere(codeSession:hString;NoeudCourant:integer):integer;
var Q:TQuery;
begin
 result:=0;
 Q:=MOPenSql('SELECT QBR_NUMNOEUDPERE FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+STRFPOINT(NoeudCourant)+'"',
             'BPFctArbre (ChercheNoeudPere)',true);
 if not Q.eof
  then result:=Q.fields[0].asInteger;
 ferme(Q);
end;

//*****************************************************************************
//Modification d'un noeud

procedure BPUpdate(const SaisiSvg,CorrectifSvg,
          saivalS,correctifvalS:hString;
          var OkModifSaisi,OkModifCorrectif,okMAJCorrectif:boolean;
          var saival,correctifval:double);
begin
 OkModifSaisi:=false;
 OkModifCorrectif:=false;
 okMAJCorrectif:=false;
 saival:=0;
 if (saivalS<>'')
  then
   begin
    saival:=Valeur(saivalS);
    OkModifSaisi:=true;
    okMAJCorrectif:=true;
   end
  else
   begin
    if (saivalS<>SaisiSvg) and (SaisiSvg<>'0')
     then
      begin
       OkModifSaisi:=true;
      end;
   end;
 correctifval:=0;
 if (correctifvalS<>'') or (correctifvalS<>CorrectifSvg)
  then
   begin
    correctifval:=Valeur(correctifvalS);
    OkModifCorrectif:=true;
   end;

 if correctifvalS=''
  then correctifval:=0;




end;



{INSERT INTO QBPARBRE 
               (QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_CODEAXE,QBR_VALEURAXE,
               QBR_LIBVALAXE,QBR_NIVEAU,QBR_DEVISE,QBR_COMMENTAIREBP,
               QBR_DATEDELAI,QBR_VALBLOQUE,QBR_VALBLOQUETMP,
               QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,
               QBR_VALAXENIV5,QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,
               QBR_VALAXENIV9,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,
               QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,
               QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_QTECPRCT,
               QBR_REF1,QBR_REFPRCT1,QBR_OP1,QBR_OPPRCT1,
               QBR_REF2,QBR_REFPRCT2,QBR_OP2,QBR_OPPRCT2,
               QBR_REF3,QBR_REFPRCT3,QBR_OP3,QBR_OPPRCT3,
               QBR_REF4,QBR_REFPRCT4,QBR_OP4,QBR_OPPRCT4,
               QBR_REF5,QBR_REFPRCT5,QBR_OP5,QBR_OPPRCT5,
               QBR_REF6,QBR_REFPRCT6,QBR_OP6,QBR_OPPRCT6,
               QBR_COEFFCALCUL,QBR_COEFFRETENU,QBR_HISTO,QBR_PREVU,QBR_REALISE,
               QBR_TOTALREF,QBR_VUREF,QBR_PERDU,QBR_RESTEAVISTER,
               QBR_TOTALCOURANT,QBR_VUCOURANT,QBR_NOUVEAU,
               QBR_PROSPECT,QBR_NEWPROSPECT,QBR_XVISIT,QBR_COEFFGENE,
               QBR_NBCLTREF,QBR_NBCLTPERDU,QBR_NBCLTNOUVEAU,
               QBR_NBCLTVU,QBR_NBCLTRESTAVOIR) 
select QBI_CODESESSION,QBI_NUMNOEUD,QBI_NUMNOEUDPERE,QBI_CODEAXE,QBI_VALEURAXE,
               QBI_LIBVALAXE,QBI_NIVEAU,QBI_DEVISE,QBI_COMMENTAIREBP,
               QBI_DATEDELAI,QBI_VALBLOQUE,QBI_VALBLOQUETMP,
               QBI_VALAXENIV1,QBI_VALAXENIV2,QBI_VALAXENIV3,QBI_VALAXENIV4,
               QBI_VALAXENIV5,QBI_VALAXENIV6,QBI_VALAXENIV7,QBI_VALAXENIV8,
               QBI_VALAXENIV9,"","","",
               "","","",
               QBI_QTEVAL,QBI_QTEPRCT,QBI_NEWQTEVAL,QBI_NEWQTEPRCT,
               QBI_HISTOVAL,QBI_HISTOPRCT,QBI_NEWVAL,QBI_NEWPRCT,
               QBI_CAVAL2,QBI_CAPRCT2,QBI_NEWCAVAL2,QBI_NEWCAPRCT2,
               QBI_CAVAL3,QBI_CAPRCT3,QBI_NEWCAVAL3,QBI_NEWCAPRCT3,
               QBI_CAVAL4,QBI_CAPRCT4,QBI_NEWCAVAL4,QBI_NEWCAPRCT4,
               QBI_CAVAL5,QBI_CAPRCT5,QBI_NEWCAVAL5,QBI_NEWCAPRCT5,
               QBI_CAVAL6,QBI_CAPRCT6,QBI_NEWCAVAL6,QBI_NEWCAPRCT6,
               "0","0","0","0","0",
               "0","0","0","0",
               "0","0","0",
               "0","0","0","0",
               "0","0","0",
               "0","0" from qbpsaisieprev}

//*****************************************************************************
//Mise à jour des % de variation

procedure MAJPrctVariation(const codeSession:hString);
var codeSql:hString;
begin
 codeSql:=' QBR_PRCTVARIATION1=IIF(QBR_REF1=0,100,(QBR_OP1/QBR_REF1)*100),'+
           'QBR_PRCTVARIATION2=IIF(QBR_REF2=0,100,(QBR_OP2/QBR_REF2)*100),'+
           'QBR_PRCTVARIATION3=IIF(QBR_REF3=0,100,(QBR_OP3/QBR_REF3)*100),'+
           'QBR_PRCTVARIATION4=IIF(QBR_REF4=0,100,(QBR_OP4/QBR_REF4)*100),'+
           'QBR_PRCTVARIATION5=IIF(QBR_REF5=0,100,(QBR_OP5/QBR_REF5)*100),'+
           'QBR_PRCTVARIATION6=IIF(QBR_REF6=0,100,(QBR_OP6/QBR_REF6)*100),'+
           'QBR_PRCTVARIATIONQ=IIF(QBR_QTEREF=0,100,(QBR_QTEC/QBR_QTEREF)*100) ';
 MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
             ' WHERE QBR_CODESESSION="'+codeSession+
             '" ',
             'BPFctArbre (MAJPrctVariation).');
end;

//
function ValeursFilsSuperieurValeurPereBloq(okPere:boolean;codeSession,noeudPere,noeud,
         codeSqlNewValeurNoeud,codeChpModif:hString;
         TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
         TotalAncCA4,TotalAncCA5,TotalAncCA6:double):boolean;
var Q:TQuery;
    sommeF,
    TotalPere,NewValNoeud:double;
    ChpSum:hString;
begin
  result:=true;
  //sommeF:=0;
  TotalPere:=0;
  if codeChpModif='1' then
  begin
    chpSum:='QBR_OP1';
    TotalPere:=TotalAncCA1;
  end;
  if codeChpModif='2' then
  begin
    chpSum:='QBR_OP2';
    TotalPere:=TotalAncCA2;
  end;
  if codeChpModif='3' then
  begin
    chpSum:='QBR_OP3';
    TotalPere:=TotalAncCA3;
  end;
  if codeChpModif='4' then
  begin
    chpSum:='QBR_OP4';
    TotalPere:=TotalAncCA4;
   end;
  if codeChpModif='5' then
  begin
    chpSum:='QBR_OP5';
    TotalPere:=TotalAncCA5;
  end;
  if codeChpModif='6' then
  begin
    chpSum:='QBR_OP6';
    TotalPere:=TotalAncCA6;
  end;
  if codeChpModif='qte' then
  begin
    chpSum:='QBR_QTEC';
    TotalPere:=TotalAncQte;
  end;

  NewValNoeud:=0;
  //nouvelle valeur du noeud
  Q:=MOpenSql('SELECT '+codeSqlNewValeurNoeud+' FROM QBPARBRE '+
              'WHERE QBR_CODESESSION="'+codeSession+
              '" AND QBR_NUMNOEUD="'+Noeud+
              '" ',//+AND QBR_NUMNOEUD<>"'+noeud+'"',
              'BPFctArbre (ValeursFilsSuperieurValeurPereBloq).',true);
  if not Q.eof then NewValNoeud:=Q.fields[0].asFloat;
  ferme(Q);

  if OkPere then
  begin
    //somme des freres bloqués
    sommeF:=0;
    Q:=MOpenSql('SELECT SUM('+chpSum+') FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUDPERE="'+noeudPere+
                '" AND QBR_VALBLOQUETMP="X"',
                'BPFctArbre (ValeursFilsSuperieurValeurPereBloq).',true);
    if not Q.eof then sommeF:=Q.fields[0].asFloat;
    ferme(Q);

    if (NewValNoeud+sommeF)-TotalPere>0 then result:=false;
  end
  else
  begin
    //somme des freres
    sommeF:=0;
    Q:=MOpenSql('SELECT SUM('+chpSum+') FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUDPERE="'+Noeud+
                '" AND QBR_VALBLOQUETMP="X"',
                'BPFctArbre (ValeursFilsSuperieurValeurPereBloq).',true);
    if not Q.eof then sommeF:=Q.fields[0].asFloat;
    ferme(Q);
    if NewValNoeud-SommeF<0 then result:=false;
  end;
end;

function OkModifNoeudArbre(codeSession,noeud,noeudPere,codeSqlNewValeurNoeud,codeChpModif:hString;
         niveau:integer;
         OkModifValBloque:boolean):boolean;
var TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3:double;
    TotalAncCA4,TotalAncCA5,TotalAncCA6:double;
    res1,res2:boolean;
begin
  res1:=true;
  res2:=true;
  MiseAJourValeurBloque(codeSession,niveau);

  //si la valeur du noeud courant n'est pas bloquée
  //ni par heritage par les fils ou le pere
  if not NoeudCourantValBloquee(codeSession,noeud,noeudPere) then
  begin
    //si la valeur du pere bloquée
    if ExisteSql('SELECT QBR_VALBLOQUE FROM QBPARBRE '+
                 'WHERE QBR_CODESESSION="'+codeSession+
                 '" AND QBR_NUMNOEUD="'+NoeudPere+
                 '" AND QBR_VALBLOQUETMP="X" ')
    then
    begin
       ArbreTotalNivRefPere(codeSession,noeudPere,
                           TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
                           TotalAncCA4,TotalAncCA5,TotalAncCA6);
       //les valeurs des fils ne doit pas être superieur à la valeur du pere
       //pour ne pas avoir de valeur négative
       res1:=ValeursFilsSuperieurValeurPereBloq(true,codeSession,noeudPere,noeud,
                           codeSqlNewValeurNoeud,codeChpModif,
                           TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
                           TotalAncCA4,TotalAncCA5,TotalAncCA6);
    end;

    //si il existe au moins un fils
    if ExisteSql('SELECT QBR_VALBLOQUE FROM QBPARBRE '+
                 'WHERE QBR_CODESESSION="'+codeSession+
                 '" AND QBR_NUMNOEUDPERE="'+Noeud+
                 '" ')
    then
    begin
      //si tous les fils sont bloques
      if not ExisteSql('SELECT QBR_VALBLOQUE FROM QBPARBRE '+
                       'WHERE QBR_CODESESSION="'+codeSession+
                       '" AND QBR_NUMNOEUDPERE="'+Noeud+
                       '" AND QBR_VALBLOQUETMP="-" ')
      then res2:=false
      else
      begin
        //si existe un fils bloqué
        if ExisteSql('SELECT QBR_VALBLOQUE FROM QBPARBRE '+
                     'WHERE QBR_CODESESSION="'+codeSession+
                     '" AND QBR_NUMNOEUDPERE="'+Noeud+
                     '" AND QBR_VALBLOQUETMP="X" ')
        then
        begin
          //la valeur du noeud doit être superieur à la somme des valeurs des fils bloqués
          //pour ne pas avoir de valeur negative
          res2:=ValeursFilsSuperieurValeurPereBloq(false,codeSession,noeudPere,noeud,
                                                   codeSqlNewValeurNoeud,codeChpModif,
                                                   TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
                                                   TotalAncCA4,TotalAncCA5,TotalAncCA6);
        end
        else res2:=true;
      end;
    end;
    result:=res1 and res2;
  end
  else result:=false;
end;

procedure CreateSousNiveauAuto(const retour,codeSession,niveau,NoeudPere:hString);
var Q:TQuery;
    ListeDates:string;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu:double;
    i,numnoeud,numnoeudDetail,NewNumNoeudOrig,j,nivDate:integer;
    DateNode:TDateTime;
    TabAxe:array [0..10] of string;
    TabEvolValues : array [1..7] of string;
begin
  ListeDates := TrouveArgument(retour,'DATESCHECKED','');
  ListeDates := AnsiReplaceText(ListeDates,'|',';');
  if ListeDates = '' then exit;

  for i:=1 to 7 do
  begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
    if i = 1 then TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIE'+IntToStr(i),'0')))
    else if i = 2 then TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIEQTE','')))
    else TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIE'+IntToStr(i-1),'')));
  end;

  numnoeud:=BPIncrementenumNoeud(codeSession);
  NewNumNoeudOrig := numnoeud;

  nivDate:=StrToInt(niveau)+1;

  Q:=OpenSql('SELECT QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
              'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10,QBR_VALEURAXE '+
             'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD='+NoeudPere,true);
  if not Q.eof then
  begin
    for j:=0 to 9 do TabAxe[j]:=Q.Fields[j].AsString;
    TabAxe[StrToInt(niveau)-1]:=Q.Fields[10].AsString;
  end;
  ferme(Q);

  While ListeDates <> '' do
  begin
    DateNode:=StrToDateTime(ReadTokenSt(ListeDates));
    ExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_NIVEAU,'+
               'QBR_DATEDELAI,QBR_CODEAXE,QBR_VALEURAXE,QBR_LIBVALAXE,'+
               'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
               'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10,'+
               'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
               'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
               'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
               'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
               'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
               'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
               'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
               'QBR_EVOLVAL1,QBR_EVOLPRCT1,QBR_SAISI1,'+
               'QBR_EVOLVAL2,QBR_EVOLPRCT2,QBR_SAISI2,'+
               'QBR_EVOLVAL3,QBR_EVOLPRCT3,QBR_SAISI3,'+
               'QBR_EVOLVAL4,QBR_EVOLPRCT4,QBR_SAISI4,'+
               'QBR_EVOLVAL5,QBR_EVOLPRCT5,QBR_SAISI5,'+
               'QBR_EVOLVAL6,QBR_EVOLPRCT6,QBR_SAISI6,'+
               'QBR_EVOLQTE,QBR_EVOLQTEPRCT,QBR_SAISIQTE,'+
               'QBR_VALBLOQUE,QBR_VALBLOQUETMP ) values ("'+codeSession+'",'+IntToStr(numnoeud)+','+NoeudPere+','+IntToStr(nivDate)+
               ',"'+USDATETIME(DateNode)+'","DELAI","'+DateTimeToStr(DateNode)+'","'+DateTimeToStr(DateNode)+'",'+
               '"'+TabAxe[0]+'","'+TabAxe[1]+'","'+TabAxe[2]+'","'+TabAxe[3]+'","'+TabAxe[4]+'",'+
               '"'+TabAxe[5]+'","'+TabAxe[6]+'","'+TabAxe[7]+'","'+TabAxe[8]+'","'+TabAxe[9]+'",'+
               TabEvolValues[1]+',0,0,0,'+
               TabEvolValues[3]+',0,0,0,'+
               TabEvolValues[4]+',0,0,0,'+
               TabEvolValues[5]+',0,0,0,'+
               TabEvolValues[6]+',0,0,0,'+
               TabEvolValues[7]+',0,0,0,'+
               TabEvolValues[2]+',0,0,0,'+
               '0,0,'+TabEvolValues[1]+','+
               '0,0,'+TabEvolValues[3]+','+
               '0,0,'+TabEvolValues[4]+','+
               '0,0,'+TabEvolValues[5]+','+
               '0,0,'+TabEvolValues[6]+','+
               '0,0,'+TabEvolValues[7]+','+
               '0,0,'+TabEvolValues[2]+',"-","-")');
    numnoeud:=numnoeud+1;
  end;

  //MAJ Arbre
  ArbreTotalNiv(codeSession,noeudpere,
                TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
  MiseAjourNiv(true,codeSession,IntToStr(numnoeud-1),noeudpere,IntToStr(nivDate),STrToInt(niveau),
               TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
               TotalNCA5,TotalNCA6,0,0,0,0,0,0,0,'0',false,false,false);

  MAJPrctVariation(codeSession);

  //MAJ Arbre detail
  numnoeudDetail:=BPIncrementeNumNoeudDetail(codeSession);
  For i:=NewNumNoeudOrig to NumNoeud-1 do
  begin
    ExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
               'QBH_DATEDELAI,'+
               'QBH_OP1,QBH_OPPRCT1,QBH_REF1,QBH_REFPRCT1,'+
               'QBH_OP2,QBH_OPPRCT2,QBH_REF2,QBH_REFPRCT2,'+
               'QBH_OP3,QBH_OPPRCT3,QBH_REF3,QBH_REFPRCT3,'+
               'QBH_OP4,QBH_OPPRCT4,QBH_REF4,QBH_REFPRCT4,'+
               'QBH_OP5,QBH_OPPRCT5,QBH_REF5,QBH_REFPRCT5,'+
               'QBH_OP6,QBH_OPPRCT6,QBH_REF6,QBH_REFPRCT6,'+
               'QBH_QTEC,QBH_QTECPRCT,QBH_QTEREF,QBH_QTEREFPRCT,'+
               'QBH_DATEPIECE,QBH_NUMNOEUDREF) SELECT QBR_CODESESSION,'+IntToStr(numnoeudDetail)+',QBR_NIVEAU,'+
               'QBR_DATEDELAI,'+
               'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
               'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
               'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
               'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
               'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
               'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
               'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
               'QBR_DATEDELAI,QBR_NUMNOEUD'+
               ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD="'+IntToStr(i)+'"');

    numnoeudDetail:=numnoeudDetail+1;
  end;
end;

procedure ModifSousNiveauAuto(const retour,codeSession,codeaxe,valeuraxe,niveau,NoeudPere:hString);
var Q,Q2:TQuery;
    i,nivMaxSession,IndiceTabReq:integer;
    ListeDateSQL,evolprct,saival,evolval,codeSql,codeaxeT:string;
    NodeBlocked,SelectBloqNodeSQL,NodeBlockedTMP,NodeBlok:string;
    noeud,noeudP,niveauP,NodeBloq,CodeNodeBloq,SelectNodeSQL,ComaSQL:string;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu:double;
    okMo,OkSessionInitPrev,OkInitNivTaille,OkModif:boolean;
    TabReqSQL : array[0..10] of string;


    function Coma(St: String): String;
    begin
      if St='' then result:='' else result:=',';
    end;

begin
  ListeDateSQL := TrouveArgument(retour,'DATESCHECKED','');
  ListeDateSQL := AnsiReplaceText(ListeDateSQL,'|','","');

  if ListeDateSQL = '' then exit;

  for i:=1 to 7 do
  begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
    if i = 1 then
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCT'+IntToStr(i),'0');
      saival:=TrouveArgument(retour,'SAISIE'+IntToStr(i),'0');
      evolval:=TrouveArgument(retour,'EVOLVAL'+IntToStr(i),'0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSQL+Coma(codeSQL)+'QBR_OP1=QBR_REF1+((QBR_REF1*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1=QBR_REF1+((QBR_REF1*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1=QBR_REF1+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end
    else if i = 2 then
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCTQTE','0');
      saival:=TrouveArgument(retour,'SAISIEQTE','0');
      evolval:=TrouveArgument(retour,'EVOLVALQTE','0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+((QBR_QTEREF*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+((QBR_QTEREF*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end
    else
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCT'+IntToStr(i-1),'0');
      saival:=TrouveArgument(retour,'SAISIE'+IntToStr(i-1),'0');
      evolval:=TrouveArgument(retour,'EVOLVAL'+IntToStr(i-1),'0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+((QBR_REF'+IntToStr(i-1)+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+((QBR_REF'+IntToStr(i-1)+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end;
  end;

  nivMaxSession:=ChercheNivMaxSession(codeSession);

  SelectNodeSQL := '';ComaSQL:='';
  IndiceTabReq:=nivMaxSession-1;
  TabReqSQL[IndiceTabReq]:= 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE = "@@NOEUDBLOQ@@"';
  For i := StrToInt(Niveau)+1 to nivMaxSession-1 do
  begin
    SelectNodeSQL := SelectNodeSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE IN(';
    ComaSQL := ComaSQL + ')';
    IndiceTabReq:=IndiceTabReq-1;
    TabReqSQL[IndiceTabReq]:= SelectNodeSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE = "@@NOEUDBLOQ@@"'+ComaSQL;
  end;
  SelectNodeSQL := SelectNodeSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE = "'+NoeudPere+'"'+ComaSQL;

  NodeBlocked:='';
  Q:=MOPenSql('SELECT QBR_NUMNOEUD,QBR_NIVEAU FROM QBPARBRE WHERE '+
               'QBR_CODESESSION="'+codeSession+'" AND QBR_VALBLOQUE="X"',
               'BPFctArbre (EvolutionSurUneValeurAxe)',true);
  InitMoveProgressForm(nil, 'Session : '+codesession , 'Veuillez patienter', Q.RecordCount,false, True);
  While not Q.Eof do
  begin
    if not MoveCurProgressForm(TraduireMemoire('Analyse des valeurs bloquées')) then exit;
    if Q.fields[1].AsInteger<nivMaxSession then
    begin
      SelectBloqNodeSQL := AnsiReplaceText(TabReqSQL[Q.fields[1].AsInteger],'@@NOEUDBLOQ@@',Q.fields[0].AsString);
      Q2:=MOPenSql('SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE '+
                   'QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD IN ('+SelectBloqNodeSQL+')',
                   'BPFctArbre (EvolutionSurUneValeurAxe)',true);
      While not Q2.Eof do
      begin
        if NodeBlocked='' then NodeBlocked:=Q2.fields[0].AsString
        else NodeBlocked:=NodeBlocked+';'+Q2.fields[0].AsString;
        Q2.Next;
      end;
      Ferme(Q2)
    end
    else
    begin
      if Q.fields[1].AsInteger=nivMaxSession then
      begin
        if NodeBlocked='' then NodeBlocked:=Q.fields[0].AsString
        else NodeBlocked:=NodeBlocked+';'+Q.fields[0].AsString;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
  FiniMoveProgressForm;

  OkSessionInitPrev:=SessionInitPrev(codeSession);
  OkInitNivTaille:=(SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession));

  if nivMaxSession=StrToInt(Niveau)
  then Q:=MOPenSql('SELECT MAX(QBR_NUMNOEUD),MAX(QBR_NUMNOEUDPERE),'+
                   'MAX(QBR_NIVEAU),COUNT(QBR_NUMNOEUD)'+
                   'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NUMNOEUDPERE="'+NoeudPere+
                   '" AND QBR_VALEURAXE IN ("'+ListeDateSql+'") GROUP BY QBR_NUMNOEUDPERE ORDER BY QBR_NUMNOEUDPERE',
                   'BPFctArbre (EvolutionSurUneValeurAxe)',true)
  else Q:=MOPenSql('SELECT MAX(QBR_NUMNOEUD),MAX(QBR_NUMNOEUDPERE),'+
                   'MAX(QBR_NIVEAU),COUNT(QBR_NUMNOEUD)'+
                   'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"AND QBR_NUMNOEUDPERE IN ('+SelectNodeSQL+
                   ') AND QBR_VALEURAXE IN ("'+ListeDateSql+'") GROUP BY QBR_NUMNOEUDPERE ORDER BY QBR_NUMNOEUDPERE',
                   'BPFctArbre (EvolutionSurUneValeurAxe)',true);
  InitMoveProgressForm(nil, 'Session : '+codesession , 'Veuillez patienter', Q.RecordCount,false, True);
  while not Q.eof do
  begin
    noeud:=Q.fields[0].asstring;
    noeudP:=Q.fields[1].asstring;
    niveauP:=Q.fields[2].asString;
    OkModif:=true;

    NodeBlockedTMP:=NodeBlocked;
    While NodeBlockedTMP<>'' do
    begin
      NodeBlok:=ReadTokenSt(NodeBlockedTMP);
      if noeudP=NodeBlok then
      begin
        OkModif:=false;
        break
      end;
    end;

    if OkModif then
    begin
        if not MoveCurProgressForm('Traitement en cours...') then exit;
        MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
                         ' WHERE QBR_CODESESSION="'+codeSession+
                         '" AND QBR_NUMNOEUDPERE ="'+NoeudP+
                         '" AND AND QBR_VALEURAXE IN ("'+ListeDateSql+'") AND QBR_VALBLOQUE="-"',
                         'BPFctArbre (EvolutionSurUneValeurAxe).');
        ArbreTotalNiv(codeSession,noeudP,
                      TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                      TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
        MiseAjourNiv(true,codeSession,Noeud,NoeudP,NiveauP,nivMaxSession,
                     TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                     TotalNCA5,TotalNCA6,0,0,0,0,0,0,0,'0',false,OkSessionInitPrev,OkInitNivTaille);

    end
    else
    begin
      NodeBloq:=NodeBloq+Coma(NodeBloq)+NoeudP;
    end;
    Q.next;

  end;
  Ferme(Q);

  MAJPrctVariation(codeSession);
  FiniMoveProgressForm;

  if NodeBloq<>'' then
  begin
    codeaxeT:=DonneCodeAxeNiv(codeSession,intToStr(nivMaxSession));
    Q:=OpenSQL('SELECT QBR_VALEURAXE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD IN ('+NodeBloq+')',true);
    While not Q.Eof do
    begin
      CodeNodeBloq:=CodeNodeBloq + #13#10 + ' - ' + DonneLibelleValeurAxe(codeaxeT,Q.Fields[0].AsString) + ' ('+Q.Fields[0].AsString+')';
      Q.next;
    end;
    PGIINFO('Les niveaux suivants sont non modifiables'+#13#10+' (blocage des valeurs des autres niveaux):' + CodeNodeBloq);
  end;
  Ferme(Q);

end;


procedure InsertNewVal(const retour,codeSession:hString);
var Q:TQuery;
    ListeDatesTmp,ListeDates,ListeAxes,nodeniveau,typesalarie,CodeNewSal,NoeudP:string;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu:double;
    NbSalarie,i,numnoeud,numnoeudDetail,NewNumNoeudOrig,j,nivDate,colSalarie,NbNewVal,nivVerif:integer;
    DateNode:TDateTime;
    TabAxe,TabAxePere:array [0..10] of string;
    TabEvolValues : array [1..7] of string;
    TabCodeAxe,TabNumValAxe:array [0..14] of hString;
    ArbreKO:boolean;
begin
  ListeDates := TrouveArgument(retour,'DATESCHECKED','');
  ListeDates := AnsiReplaceText(ListeDates,'|',';');
  if ListeDates = '' then exit;

  for i:=1 to 7 do
  begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
    if i = 1 then TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIE'+IntToStr(i),'0')))
    else if i = 2 then TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIEQTE','')))
    else TabEvolValues[i]:=STRFPOINT(VALEUR(TrouveArgument(retour,'SAISIE'+IntToStr(i-1),'')));
  end;

  numnoeud:=BPIncrementenumNoeud(codeSession);
  NewNumNoeudOrig := numnoeud;

  j:=0;

  nivDate:=1;
  colSalarie:=1;
  ListeAxes:=TrouveArgument(retour,'VALEURSAXES','');
  ListeAxes := AnsiReplaceText(ListeAxes,'|',';');
  While ListeAxes<>'' do
  begin
    TabAxe[j]:=ReadTokenSt(ListeAxes);
    if TabAxe[j]='VIDE' then TabAxe[j]:='';
    if TabAxe[j]='DATES' then
    begin
      nivDate:=j+1;
      nodeniveau:=IntTostr(j);
      TabAxe[j]:='';
    end;
    if TabAxe[j]='SALARIE' then
    begin
      colSalarie:=j;
      TabAxe[j]:='';
    end;
    j:=j+1;
  end;

  NbSalarie:=StrToInt(TrouveArgument(retour,'NBNEWVAL',''));
  TypeSalarie:=TrouveArgument(retour,'TYPESAL','');

  NbNewVal:=0;
  Q := OpenSQL('SELECT QBR_VALEURAXE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" '+
               'AND QBR_VALEURAXE LIKE "'+TypeSalarie+'%" ORDER BY QBR_VALEURAXE DESC', True, 1);
  if not Q.eof then
  begin
    NbNewVal := StrToInt(Copy(Q.fields[0].asString,3,8))
  end;
  ferme(Q);

  ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);
  nivVerif:=0;
  NoeudP:='0';
  ArbreKO:=true;
  While ArbreKO do
  begin
    TabAxePere[0]:=TabAxe[nivVerif];
    for j:=1 to nivVerif do TabAxePere[j]:=TabAxe[j-1];
    for j:=nivVerif+1 to 10 do TabAxePere[j]:='';

    //Vérification existence arborescence
    if not ExisteSQL('SELECT 1 FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                     '" AND QBR_VALEURAXE="'+TabAxePere[0]+'" AND QBR_VALAXENIV1="'+TabAxePere[1]+
                     '" AND QBR_VALAXENIV2="'+TabAxePere[2]+'" AND QBR_VALAXENIV3="'+TabAxePere[3]+
                     '" AND QBR_VALAXENIV4="'+TabAxePere[4]+'" AND QBR_VALAXENIV5="'+TabAxePere[5]+
                     '" AND QBR_VALAXENIV6="'+TabAxePere[6]+'" AND QBR_VALAXENIV7="'+TabAxePere[7]+
                     '" AND QBR_VALAXENIV8="'+TabAxePere[8]+'" AND QBR_VALAXENIV9="'+TabAxePere[9]+'"') then
    begin
      ExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_NIVEAU,'+
                 'QBR_DATEDELAI,QBR_CODEAXE,QBR_VALEURAXE,QBR_LIBVALAXE,'+
                 'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
                 'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10,'+
                 'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                 'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                 'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                 'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                 'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                 'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                 'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                 'QBR_EVOLVAL1,QBR_EVOLPRCT1,QBR_SAISI1,'+
                 'QBR_EVOLVAL2,QBR_EVOLPRCT2,QBR_SAISI2,'+
                 'QBR_EVOLVAL3,QBR_EVOLPRCT3,QBR_SAISI3,'+
                 'QBR_EVOLVAL4,QBR_EVOLPRCT4,QBR_SAISI4,'+
                 'QBR_EVOLVAL5,QBR_EVOLPRCT5,QBR_SAISI5,'+
                 'QBR_EVOLVAL6,QBR_EVOLPRCT6,QBR_SAISI6,'+
                 'QBR_EVOLQTE,QBR_EVOLQTEPRCT,QBR_SAISIQTE,'+
                 'QBR_VALBLOQUE,QBR_VALBLOQUETMP ) values ("'+codeSession+'",'+IntToStr(numnoeud)+','+NoeudP+','+IntToStr(nivVerif)+
                 ',"'+USDATETIME(idate1900)+'","'+TabCodeAxe[NivVerif+1]+'","'+TabAxePere[0]+'","",'+
                 '"'+TabAxePere[1]+'","'+TabAxePere[2]+'","'+TabAxePere[3]+'","'+TabAxePere[4]+'","'+TabAxePere[5]+'",'+
                 '"'+TabAxePere[6]+'","'+TabAxePere[7]+'","'+TabAxePere[8]+'","'+TabAxePere[9]+'","",'+
                 '0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+
                 '0,0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+
                 '0,0,0,"-","-")');

      NoeudP:=IntToStr(numnoeud);
      numnoeud:=numnoeud+1;
      nivVerif:=nivVerif+1;
      if nivVerif=colSalarie then ArbreKO:=false;
    end
    else
    begin
      Q:=OpenSQL('SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                 '" AND QBR_VALEURAXE="'+TabAxePere[0]+'" AND QBR_VALAXENIV1="'+TabAxePere[1]+
                 '" AND QBR_VALAXENIV2="'+TabAxePere[2]+'" AND QBR_VALAXENIV3="'+TabAxePere[3]+
                 '" AND QBR_VALAXENIV4="'+TabAxePere[4]+'" AND QBR_VALAXENIV5="'+TabAxePere[5]+
                 '" AND QBR_VALAXENIV6="'+TabAxePere[6]+'" AND QBR_VALAXENIV7="'+TabAxePere[7]+
                 '" AND QBR_VALAXENIV8="'+TabAxePere[8]+'" AND QBR_VALAXENIV9="'+TabAxePere[9]+'"',true);
      if not Q.eof then NoeudP:=Q.fields[0].asString;
      ferme(Q);

      nivVerif:=nivVerif+1;
      if nivVerif=colSalarie then ArbreKO:=false;
    end
  end;

  for i:=1 to NbSalarie do
  begin
    TabAxePere[0]:=TabAxe[colSalarie-1];
    for j:=1 to 9 do
    begin
      if colSalarie-1-j>=0 then TabAxePere[j]:=TabAxe[colSalarie-1-j]
      else TabAxePere[j]:='';
    end;

    Q:=OpenSQL('SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_VALEURAXE="'+TabAxePere[0]+'" AND QBR_VALAXENIV1="'+TabAxePere[1]+
               '" AND QBR_VALAXENIV2="'+TabAxePere[2]+'" AND QBR_VALAXENIV3="'+TabAxePere[3]+
               '" AND QBR_VALAXENIV4="'+TabAxePere[4]+'" AND QBR_VALAXENIV5="'+TabAxePere[5]+
               '" AND QBR_VALAXENIV6="'+TabAxePere[6]+'" AND QBR_VALAXENIV7="'+TabAxePere[7]+
               '" AND QBR_VALAXENIV8="'+TabAxePere[8]+'" AND QBR_VALAXENIV9="'+TabAxePere[9]+'"',true);
    if not Q.eof then NoeudP:=Q.fields[0].asString;
    ferme(Q);

    TabAxe[colSalarie]:='';
    CodeNewSal:=TypeSalarie+MetZero(IntToStr(NbNewVal+i),8);
    ExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_NIVEAU,'+
               'QBR_DATEDELAI,QBR_CODEAXE,QBR_VALEURAXE,QBR_LIBVALAXE,'+
               'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
               'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10,'+
               'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
               'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
               'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
               'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
               'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
               'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
               'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
               'QBR_EVOLVAL1,QBR_EVOLPRCT1,QBR_SAISI1,'+
               'QBR_EVOLVAL2,QBR_EVOLPRCT2,QBR_SAISI2,'+
               'QBR_EVOLVAL3,QBR_EVOLPRCT3,QBR_SAISI3,'+
               'QBR_EVOLVAL4,QBR_EVOLPRCT4,QBR_SAISI4,'+
               'QBR_EVOLVAL5,QBR_EVOLPRCT5,QBR_SAISI5,'+
               'QBR_EVOLVAL6,QBR_EVOLPRCT6,QBR_SAISI6,'+
               'QBR_EVOLQTE,QBR_EVOLQTEPRCT,QBR_SAISIQTE,'+
               'QBR_VALBLOQUE,QBR_VALBLOQUETMP ) values ("'+codeSession+'",'+IntToStr(numnoeud)+','+NoeudP+','+IntToStr(colSalarie+1)+
               ',"'+USDATETIME(idate1900)+'","011","'+CodeNewSal+'","",'+
               '"'+TabAxe[0]+'","'+TabAxe[1]+'","'+TabAxe[2]+'","'+TabAxe[3]+'","'+TabAxe[4]+'",'+
               '"'+TabAxe[5]+'","'+TabAxe[6]+'","'+TabAxe[7]+'","'+TabAxe[8]+'","'+TabAxe[9]+'",'+
               '0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+'0,0,0,0,'+
               '0,0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+'0,0,0,'+
               '0,0,0,"-","-")');

    NoeudP:=IntToStr(numnoeud);
    numnoeud:=numnoeud+1;
    TabAxe[colSalarie]:=TypeSalarie+MetZero(IntToStr(NbNewVal+i),8);
    ListeDatesTmp:=ListeDates;
    While ListeDatesTmp <> '' do
    begin
      DateNode:=StrToDateTime(ReadTokenSt(ListeDatesTmp));
      ExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_NIVEAU,'+
                 'QBR_DATEDELAI,QBR_CODEAXE,QBR_VALEURAXE,QBR_LIBVALAXE,'+
                 'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,'+
                 'QBR_VALAXENIV6,QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10,'+
                 'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                 'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                 'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                 'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                 'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                 'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                 'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                 'QBR_EVOLVAL1,QBR_EVOLPRCT1,QBR_SAISI1,'+
                 'QBR_EVOLVAL2,QBR_EVOLPRCT2,QBR_SAISI2,'+
                 'QBR_EVOLVAL3,QBR_EVOLPRCT3,QBR_SAISI3,'+
                 'QBR_EVOLVAL4,QBR_EVOLPRCT4,QBR_SAISI4,'+
                 'QBR_EVOLVAL5,QBR_EVOLPRCT5,QBR_SAISI5,'+
                 'QBR_EVOLVAL6,QBR_EVOLPRCT6,QBR_SAISI6,'+
                 'QBR_EVOLQTE,QBR_EVOLQTEPRCT,QBR_SAISIQTE,'+
                 'QBR_VALBLOQUE,QBR_VALBLOQUETMP ) values ("'+codeSession+'",'+IntToStr(numnoeud)+','+NoeudP+','+IntToStr(nivDate)+
                 ',"'+USDATETIME(DateNode)+'","DELAI","'+DateTimeToStr(DateNode)+'","'+DateTimeToStr(DateNode)+'",'+
                 '"'+TabAxe[0]+'","'+TabAxe[1]+'","'+TabAxe[2]+'","'+TabAxe[3]+'","'+TabAxe[4]+'",'+
                 '"'+TabAxe[5]+'","'+TabAxe[6]+'","'+TabAxe[7]+'","'+TabAxe[8]+'","'+TabAxe[9]+'",'+
                 TabEvolValues[1]+',0,0,0,'+
                 TabEvolValues[3]+',0,0,0,'+
                 TabEvolValues[4]+',0,0,0,'+
                 TabEvolValues[5]+',0,0,0,'+
                 TabEvolValues[6]+',0,0,0,'+
                 TabEvolValues[7]+',0,0,0,'+
                 TabEvolValues[2]+',0,0,0,'+
                 '0,0,'+TabEvolValues[1]+','+
                 '0,0,'+TabEvolValues[3]+','+
                 '0,0,'+TabEvolValues[4]+','+
                 '0,0,'+TabEvolValues[5]+','+
                 '0,0,'+TabEvolValues[6]+','+
                 '0,0,'+TabEvolValues[7]+','+
                 '0,0,'+TabEvolValues[2]+',"-","-")');
      numnoeud:=numnoeud+1;
    end;

    //MAJ Arbre
    //TEST SIC
    //ArbreTotalNiv(codeSession,noeudpere,
    ArbreTotalNiv(codeSession,'',TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                  TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
    MiseAjourNiv(true,codeSession,IntToStr(numnoeud-1),NoeudP,IntToStr(nivDate),STrToInt(nodeniveau),
                 TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                 TotalNCA5,TotalNCA6,0,0,0,0,0,0,0,'0',false,false,false);
     MAJPrctVariation(codeSession);
   end;

  //MAJ Arbre detail
  numnoeudDetail:=BPIncrementeNumNoeudDetail(codeSession);
  For i:=NewNumNoeudOrig to NumNoeud-1 do
  begin
    ExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
               'QBH_DATEDELAI,'+
               'QBH_OP1,QBH_OPPRCT1,QBH_REF1,QBH_REFPRCT1,'+
               'QBH_OP2,QBH_OPPRCT2,QBH_REF2,QBH_REFPRCT2,'+
               'QBH_OP3,QBH_OPPRCT3,QBH_REF3,QBH_REFPRCT3,'+
               'QBH_OP4,QBH_OPPRCT4,QBH_REF4,QBH_REFPRCT4,'+
               'QBH_OP5,QBH_OPPRCT5,QBH_REF5,QBH_REFPRCT5,'+
               'QBH_OP6,QBH_OPPRCT6,QBH_REF6,QBH_REFPRCT6,'+
               'QBH_QTEC,QBH_QTECPRCT,QBH_QTEREF,QBH_QTEREFPRCT,'+
               'QBH_DATEPIECE,QBH_NUMNOEUDREF) SELECT QBR_CODESESSION,'+IntToStr(numnoeudDetail)+',QBR_NIVEAU,'+
               'QBR_DATEDELAI,'+
               'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
               'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
               'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
               'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
               'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
               'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
               'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
               'QBR_DATEDELAI,QBR_NUMNOEUD'+
               ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD="'+IntToStr(i)+'"');

    numnoeudDetail:=numnoeudDetail+1;
  end;
end;


procedure EvolutionTwo(const retour,codeSession:hString);
var Q,Q2:TQuery;
    i,nivMaxSession:integer;
    ListeDateSQL,evolprct,saival,evolval,codeSql,codeaxeT:string;
    NodeBlocked,SelectBloqNodeSQL,NodeBlockedTMP,NodeBlok:string;
    noeud,noeudP,niveauP,NodeBloq,CodeNodeBloq,SelectNodeSQL,ComaSQL,ListeValues,ValueAxe:string;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu:double;
    okMo,OkSessionInitPrev,OkInitNivTaille,OkModif:boolean;
    TabReqSQL : array[0..10] of string;
    ReqValAxe:string;

    function Coma(St: String): String;
    begin
      if St='' then result:='' else result:=',';
    end;

begin
  ListeDateSQL := TrouveArgument(retour,'DATESCHECKED','');
  ListeDateSQL := AnsiReplaceText(ListeDateSQL,'|','","');

  if ListeDateSQL = '' then exit;

  ListeValues := TrouveArgument(retour,'VALEURSAXES','');
  ListeValues := AnsiReplaceText(ListeValues,'|',';');

  i:=1;
  While ListeValues<>'' do
  begin
    ValueAxe:=ReadTokenSt(ListeValues);
    if ValueAxe<>'VIDE' then
    begin
      if i=1 then ReqValAxe:=' QBR_VALAXENIV'+IntToStr(i)+' IN ("'+ValueAxe
      else ReqValAxe:=ReqValAxe+'") AND QBR_VALAXENIV'+IntToStr(i)+' IN("'+ValueAxe
    end;
    i:=i+1;
  end;
  ReqValAxe:=ReqValAxe+'")';

  for i:=1 to 7 do
  begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
    if i = 1 then
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCT'+IntToStr(i),'0');
      saival:=TrouveArgument(retour,'SAISIE'+IntToStr(i),'0');
      evolval:=TrouveArgument(retour,'EVOLVAL'+IntToStr(i),'0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSQL+Coma(codeSQL)+'QBR_OP1=QBR_REF1+((QBR_REF1*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1=QBR_REF1+((QBR_REF1*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP1=QBR_REF1+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end
    else if i = 2 then
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCTQTE','0');
      saival:=TrouveArgument(retour,'SAISIEQTE','0');
      evolval:=TrouveArgument(retour,'EVOLVALQTE','0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+((QBR_QTEREF*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+((QBR_QTEREF*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_QTEC=QBR_QTEREF+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end
    else
    begin
      evolprct:=TrouveArgument(retour,'EVOLPRCT'+IntToStr(i-1),'0');
      saival:=TrouveArgument(retour,'SAISIE'+IntToStr(i-1),'0');
      evolval:=TrouveArgument(retour,'EVOLVAL'+IntToStr(i-1),'0');
      okMo:=false;
      if saival<>'' then codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'="'+STRFPOINT(VALEUR(saival))+'" '
      else
      begin
        if (evolprct<>'') and (evolval<>'') then
        begin
          codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+((QBR_REF'+IntToStr(i-1)+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) + "'+STRFPOINT(VALEUR(evolval))+'" ';
          okMo:=true;
        end;
        if (evolprct<>'') and (evolval='') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+((QBR_REF'+IntToStr(i-1)+'*"'+STRFPOINT(VALEUR(evolprct))+'")/100) ';
        if (okMo=false) and (evolval<>'') then
        codeSql:=codeSql+Coma(codeSQL)+'QBR_OP'+IntToStr(i-1)+'=QBR_REF'+IntToStr(i-1)+'+ "'+STRFPOINT(VALEUR(evolval))+'" ';
      end;
    end;
  end;

  nivMaxSession:=ChercheNivMaxSession(codeSession);

  SelectNodeSQL := '';ComaSQL:='';

  TabReqSQL[nivMaxSession-1]:= 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE = "@@NOEUDBLOQ@@"';
  For i := nivMaxSession-2  downto 1 do
  begin
    SelectNodeSQL := SelectNodeSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE IN(';
    ComaSQL := ComaSQL + ')';
    TabReqSQL[i]:= SelectNodeSQL + 'SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE QBR_NUMNOEUDPERE = "@@NOEUDBLOQ@@"'+ComaSQL;
  end;

  NodeBlocked:='';
  Q:=MOPenSql('SELECT QBR_NUMNOEUD,QBR_NIVEAU FROM QBPARBRE WHERE '+
               'QBR_CODESESSION="'+codeSession+'" AND QBR_VALBLOQUE="X"',
               'BPFctArbre (EvolutionTwo)',true);
  InitMoveProgressForm(nil, 'Session : '+codesession , 'Veuillez patienter', Q.RecordCount,false, True);
  While not Q.Eof do
  begin
    if not MoveCurProgressForm(TraduireMemoire('Analyse des valeurs bloquées')) then exit;
    if Q.fields[1].AsInteger<nivMaxSession then
    begin
      SelectBloqNodeSQL := AnsiReplaceText(TabReqSQL[Q.fields[1].AsInteger],'@@NOEUDBLOQ@@',Q.fields[0].AsString);
      Q2:=MOPenSql('SELECT QBR_NUMNOEUD FROM QBPARBRE WHERE '+
                   'QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD IN ('+SelectBloqNodeSQL+')',
                   'BPFctArbre (EvolutionTwo)',true);
      While not Q2.Eof do
      begin
        if NodeBlocked='' then NodeBlocked:=Q2.fields[0].AsString
        else NodeBlocked:=NodeBlocked+';'+Q2.fields[0].AsString;
        Q2.Next;
      end;
      Ferme(Q2)
    end
    else
    begin
      if Q.fields[1].AsInteger=nivMaxSession then
      begin
        if NodeBlocked='' then NodeBlocked:=Q.fields[0].AsString
        else NodeBlocked:=NodeBlocked+';'+Q.fields[0].AsString;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
  FiniMoveProgressForm;

  OkSessionInitPrev:=SessionInitPrev(codeSession);
  OkInitNivTaille:=(SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession));

  Q:=MOPenSql('SELECT MAX(QBR_NUMNOEUD),MAX(QBR_NUMNOEUDPERE),'+
              'MAX(QBR_NIVEAU),COUNT(QBR_NUMNOEUD)'+
              'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"AND '+ReqValAxe+
              ' AND QBR_VALEURAXE IN ("'+ListeDateSql+'") GROUP BY QBR_NUMNOEUDPERE ORDER BY QBR_NUMNOEUDPERE',
              'BPFctArbre (EvolutionTwo)',true);
  InitMoveProgressForm(nil, 'Session : '+codesession , 'Veuillez patienter', Q.RecordCount,false, True);
  while not Q.eof do
  begin
    noeud:=Q.fields[0].asstring;
    noeudP:=Q.fields[1].asstring;
    niveauP:=Q.fields[2].asString;
    OkModif:=true;

    NodeBlockedTMP:=NodeBlocked;
    While NodeBlockedTMP<>'' do
    begin
      NodeBlok:=ReadTokenSt(NodeBlockedTMP);
      if noeudP=NodeBlok then
      begin
        OkModif:=false;
        break
      end;
    end;

    if OkModif then
    begin
        if not MoveCurProgressForm('Traitement en cours...') then exit;
        MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
                         ' WHERE QBR_CODESESSION="'+codeSession+
                         '" AND QBR_NUMNOEUDPERE ="'+NoeudP+
                         '" AND AND QBR_VALEURAXE IN ("'+ListeDateSql+'") AND QBR_VALBLOQUE="-"',
                         'BPFctArbre (EvolutionTwo).');
        ArbreTotalNiv(codeSession,noeudP,
                      TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                      TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
        MiseAjourNiv(true,codeSession,Noeud,NoeudP,NiveauP,nivMaxSession,
                     TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                     TotalNCA5,TotalNCA6,0,0,0,0,0,0,0,'0',false,OkSessionInitPrev,OkInitNivTaille);

    end
    else
    begin
      NodeBloq:=NodeBloq+Coma(NodeBloq)+NoeudP;
    end;
    Q.next;

  end;
  Ferme(Q);

  MAJPrctVariation(codeSession);
  FiniMoveProgressForm;

  if NodeBloq<>'' then
  begin
    codeaxeT:=DonneCodeAxeNiv(codeSession,intToStr(nivMaxSession));
    Q:=OpenSQL('SELECT QBR_VALEURAXE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD IN ('+NodeBloq+')',true);
    While not Q.Eof do
    begin
      CodeNodeBloq:=CodeNodeBloq + #13#10 + ' - ' + DonneLibelleValeurAxe(codeaxeT,Q.Fields[0].AsString) + ' ('+Q.Fields[0].AsString+')';
      Q.next;
    end;
    PGIINFO('Les niveaux suivants sont non modifiables'+#13#10+' (blocage des valeurs des autres niveaux):' + CodeNodeBloq);
  end;
  Ferme(Q);

end;


end.
