unit BPFctSession;

interface

uses HEnt1,classes;

function ChercheNivMaxSession(const codeSession:hString):integer;
function ChercheNivMax(const codeSession:hString):integer ;
function ChercheCodeRestrictionSession(const codeAs,codeChamp,prefixe,session:hString;
         var CodeJoin:hString):hString;
function ChercheSessionStructure(const codeSession:hString):hString;
procedure ChercheCodeAxeSession(codeSession:hString;
          var TabCodeAxe:array of hString);
procedure ChercheTabCodeAxeTabNumValAxe(const codeSession:hString;
          var TabCodeAxe,TabNumValAxe:array of hString);
procedure ChercheDixNumValAxe(const structure:hString;
          TabCodeAxe:array of hString;
          var TabNumValAxe:array of hString);
procedure ChercheCodeAxeSessionTStL(codeSession:hString;
          TStCodeAxe:TStringList);

function  ChercheSessionObjectif(const session:hString;
          var SessionObjectif:hString):boolean;

//période et date
procedure DonneDateDebEtFinSaison(const Saison:hString;
          var dateDeb,dateFin:TDateTime);
procedure ChercheDateDDateFPeriode(const session:hString;
          var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime);
procedure ChercheSaisonCmd(const session:hString;
          var SaiCmd,SaiCmdRef:hString);

//ok session
function OkSessionObjectif(const codesession:hString):boolean;
function OkSessionModif(const codeSession:hString):boolean;
function OkSessionPermanent(const codeSession:hString):boolean;    
function OkSessionModifVue(const codeSession:hString):boolean;

//
function DonneValeurAffiche(const codeSession:hString):integer;

//etat session
function SessionDelai(const session:hString):boolean;
function SessionValide(const session:hString):boolean;
procedure SessionMAJ(const session:hString);
function SessionCalculParDelai(const session:hString;
          var BPInitialise:hString):boolean;            
function SessionCalculParTaille(const codeSession:hString):boolean;
function SessionInitPrev(const session:hString):boolean;     
function SessionInitCoeff(const session:hString):boolean;    
function SessionEclateeParDelai(const session:hString):boolean;
function SessionEclateeParTaille(const session:hString):boolean;
function SessionMaillageJour(const codeSession:hString):boolean;

//modif session
procedure ModifSession(const codeSession,Modif:hString);
procedure ModifSessionVue(const codeSession,Modif,ModifChpSession:hString);

function ChercheClientArbreSession(const codeSession:hString;
         nivMax:integer):hString;

//duplication
procedure DuplicationArbreBP(const codeSession,codeSessionN:hString);
procedure DuplicationArbrePartielleBP(const codeSession,codeSessionN,Niv:hString);

//methode
function  DonneMethodeSession(const codeSession:hString):hString;

//Cherche numero session
function ChercheNumSession(const codeSession:hString):hString;
//controle
function ControleSession(session: hString): boolean;

function IncrementeCodeSession(const codeSession:hString):hString;

procedure RechercheCodeRestrictionMultiValeurPGI(const codeAs,codeAsJ,codeaxe,valaxe:hString;
          var code,CodeJoin:hString);

function DonneObjectifsEtablissements(codeSession:hString;
         LesEtablissements:TStringList;DateDebut,DateFin:TDateTime;
         var LibelleValeur,Devise:hString):double;

function existCodeAxe(const TabCodeAxe:array of hString; const CodeAxe: hString): boolean;

implementation

uses Sysutils,hmsgbox,controls,HCtrls,
     {$IFNDEF EAGLCLIENT}{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else} uTob, {$ENDIF}
     BPBasic,UCtx,UUtil,SynScriptBP,BPUtil;

//cherche le nombre de niveau definit dans la session donnée
function ChercheNivMax(const codeSession:hString):integer ;
var Q:TQuery;
    nivMax,i:integer;
begin
  nivMax:=0;
  Q:=MopenSql('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4'+
              ',QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,'+
              'QBS_CODEAXES9,QBS_CODEAXES10 '+
              'FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+'"',
              'BPFctSession (ChercheNivMax).',true);
  i:= 9;
  while not Q.eof do
  begin
  if Q.fields[i].asString<>'' then
     begin
          nivmax := i+1;
          break;
     end
  else i := i-1;
  end;

  ferme(Q);
 result:=nivMax;
end;

function ChercheNivMaxSession(const codeSession:hString):integer;
var nivMax:integer;
begin
 nivMax:=ChercheNivMax(codeSession); 
 result:=NivMax;
 if not SessionInitPrev(codeSession)
  then exit;
 //niv article
 nivMax:=nivMax+1;
 //niv coloris
 if GestionBPColoris
  then nivMax:=nivMax+1;
 //niv FS
 if GestionBPFS
  then nivMax:=nivMax+1;
 //niv Magasin
 if GestionBPMagasin
  then nivMax:=nivMax+1;
 //niv delai
 if SessionDelai(codeSession)
  then nivMax:=nivMax+1;
 result:=NivMax;
end;


procedure RechercheCodeRestrictionMultiValeurPGI(const codeAs,codeAsJ,codeaxe,valaxe:hString;
          var code,CodeJoin:hString);
var chp,S,join,D:hString;
    ii,i,L:integer;
begin
 DonneCodeChpJoinPGI(codeaxe,codeAsJ,chp,join);
 CodeJoin:=CodeJoin+' '+join;
 //multi valeur possible pour valeur axe
 S:=valAxe;
 ii:=0;
 while pos(';',S)<>0 do
  begin
   i:=pos(';',S);
   L:=length(S);
   D:=copy(S,0,i-1);
   S:=copy(S,i+1,L-i);
   if ii=0
    then code:=code+' AND ('+codeAs+chp+'="'+D+'" '
    else code:=code+' OR '+codeAs+chp+'="'+D+'" ';
   ii:=ii+1;
  end;

 if ii=0
  then
   begin
    if (S<>'') and (pos(Uppercase(TraduireMemoire('Tous')),Uppercase(S))=0)
     then code:=code+' AND ('+codeAs+chp+'="'+S+'") ';
   end
  else
   begin
    if (S<>'') and (pos(Uppercase(TraduireMemoire('Tous')),Uppercase(S))=0)
     then code:=code+' OR '+codeAs+chp+'="'+S+'") '
     else code:=code+') ';
   end;
end;

function RechercheCodeRestrictionPGI(codeAs,natureCmd,qualifpiece,axe1,axe2,axe3,axe4,
         valaxe1,valaxe2,valaxe3,valaxe4:hString;
         var CodeJoin:hString):hString;
var code,D:hstring;
    ii,i,L:integer;

      function MiseEnForme(codechamp,S:string):string;
      begin
        //S:=naturecmd;
        if pos(UPPERCASE(TraduireMemoire('Tous')),UPPERCASE(S))<>0 then
        begin
          result:=code;
          exit;
        end;
        ii:=0;
        while pos(';',S)<>0 do
        begin
          i:=pos(';',S);
          L:=length(S);
          D:=copy(S,0,i-1);
          S:=copy(S,i+1,L-i);
          if ii=0 then code:=code+' AND ('+codeAs+codeChamp+'="'+D+'" '
          else code:=code+' OR '+codeAs+codeChamp+'="'+D+'" ';
          ii:=ii+1;
        end;

        if ii=0 then
        begin
          if S<>'' then code:=code+' AND ('+codeAs+codeChamp+'="'+S+'" ) ';
        end
        else
        begin
          if S<>'' then code:=code+' OR '+codeAs+codeChamp+'="'+S+'") '
          else code:=code+') ';
        end;
        result:=code;
        code:='';
      end;

begin
 //cherche le code sql
 code:='';
 if axe1<>'' then RechercheCodeRestrictionMultiValeurPGI(codeAs,'Y1',axe1,valaxe1,code,CodeJoin);
 if axe2<>'' then RechercheCodeRestrictionMultiValeurPGI(codeAs,'Y2',axe2,valaxe2,code,CodeJoin);
 if axe3<>'' then RechercheCodeRestrictionMultiValeurPGI(codeAs,'Y3',axe3,valaxe3,code,CodeJoin);
 if axe4<>'' then RechercheCodeRestrictionMultiValeurPGI(codeAs,'Y4',axe4,valaxe4,code,CodeJoin);

 Case ContextBP of
   0,1 : result := MiseEnForme('GL_NATUREPIECEG',naturecmd);
   2 : result := MiseEnForme('Y_NATUREPIECE',naturecmd) + MiseEnForme('Y_QUALIFPIECE',qualifpiece);
   3 : result := code;
 end;

end;

procedure RechercheCodeRestrictionMultiValeurAutre(const codeChamp,valeurAxe,num:hString;
          const codeAs:hString;
          var code:hString);
var S,D:hString;
    ii,i,L:integer;
begin
 //multi valeur possible pour valeur axe
 S:=valeurAxe;
 if pos(Uppercase(TraduireMemoire('Tous')),Uppercase(S))<>0
  then exit;
 ii:=0;
 while pos(';',S)<>0 do
  begin
   i:=pos(';',S);
   L:=length(S);
   D:=copy(S,0,i-1);
   S:=copy(S,i+1,L-i);
   if ii=0
//    then code:=code+' AND ('+codeAs+codeChamp+num+' LIKE "%'+D+'%" '
    then code:=code+' AND ('+codeAs+codeChamp+num+' = "'+D+'" '
//    else code:=code+' OR '+codeAs+codeChamp+num+' LIKE "%'+D+'%" ';
    else code:=code+' OR '+codeAs+codeChamp+num+' = "'+D+'" ';
   ii:=ii+1;
  end;

 if ii=0
  then
   begin
    if S<>''
//     then code:=code+' AND ('+codeAs+codeChamp+num+' LIKE "%'+S+'%" ) ';
     then code:=code+' AND ('+codeAs+codeChamp+num+' = "'+S+'" ) ';
   end
  else
   begin
    if S<>''
//     then code:=code+' OR '+codeAs+codeChamp+num+' LIKE "%'+S+'%") '
     then code:=code+' OR '+codeAs+codeChamp+num+' = "'+S+'") '
     else code:=code+') ';
   end;
end;

//cherche dans la structure de la table pivot
//où se trouvent les codes choisis
//pour la restriction au niveau de la session
procedure ChercheNumValAxeRestriction(const structure:hString;
          axe1,axe2,axe3,axe4:hString;
          var num1,num2,num3,num4:hString);
var Q:TQuery;
    i:integer;
begin
 num1:='';
 num2:='';
 num3:='';
 num4:='';
 Q:=MOpenSql('SELECT QBC_CODESTRUCT,QBC_CODEAXES1,QBC_CODEAXES2,'+
             'QBC_CODEAXES3,QBC_CODEAXES4,QBC_CODEAXES5,QBC_CODEAXES6,'+
             'QBC_CODEAXES7,QBC_CODEAXES8,QBC_CODEAXES9,QBC_CODEAXES10,'+
             'QBC_CODEAXES11,QBC_CODEAXES12,QBC_CODEAXES13,'+
             'QBC_CODEAXES14,QBC_CODEAXES15,QBC_CODEAXES16,'+
             'QBC_CODEAXES17,QBC_CODEAXES18,QBC_CODEAXES19,'+
             'QBC_CODEAXES20,QBC_CODEAXES21,QBC_CODEAXES22,'+
             'QBC_CODEAXES23,QBC_CODEAXES24,QBC_CODEAXES25,'+
             'QBC_CODEAXES26,QBC_CODEAXES27,QBC_CODEAXES28,'+
             'QBC_CODEAXES29,QBC_CODEAXES30 '+
             'FROM QBPSTRUCTURE WHERE QBC_CODESTRUCT="'+structure+'"',
             'BPFctSession (ChercheNumValAxeRestriction).',true);
 if not Q.eof
  then
   begin
    for i:=1 to 30 do
     begin
      if (Q.Fields[i].asString=axe1) and (axe1<>'')
       then num1:=IntToStr(i);
      if (Q.Fields[i].asString=axe2) and (axe2<>'')
       then num2:=IntToStr(i);
      if (Q.Fields[i].asString=axe3) and (axe3<>'')
       then num3:=IntToStr(i);
      if (Q.Fields[i].asString=axe4) and (axe4<>'')
       then num4:=IntToStr(i);
     end;
   end;
 ferme(Q);
end;

function RechercheCodeRestrictionAutre(codeAs,codeChamp,prefixe,structure,naturecmd,
         axe1,axe2,axe3,axe4,
         valaxe1,valaxe2,valaxe3,valaxe4:hString):hString;
var code,num1,num2,num3,num4:hString;
begin
 //cherche num correspondant dans pivot
 ChercheNumValAxeRestriction(structure,axe1,axe2,axe3,axe4,
                             num1,num2,num3,num4);
 //cherche le code sql
 code:='';
 if num1<>''
  then RechercheCodeRestrictionMultiValeurAutre(codeChamp,valaxe1,num1,codeAs,code);
 if num2<>''
  then RechercheCodeRestrictionMultiValeurAutre(codeChamp,valaxe2,num2,codeAs,code);
 if num3<>''
  then RechercheCodeRestrictionMultiValeurAutre(codeChamp,valaxe3,num3,codeAs,code);
 if num4<>''
  then RechercheCodeRestrictionMultiValeurAutre(codeChamp,valaxe4,num4,codeAs,code);
 if (naturecmd<>'') and (naturecmd<>'%')
  then code:=code+' AND '+codeAs+Prefixe+'NATURECMD="'+naturecmd+'" ';
 result:=code;
end;

//cherche le code sql pour la restriction
function ChercheCodeRestrictionSession(const codeAs,codeChamp,prefixe,session:hString;
         var CodeJoin:hString):hString;
var structure,code,valRestriction:hString;
    Q:TQuery;
    axe1,axe2,axe3,axe4:hString;
    valaxe1,valaxe2,valaxe3,valaxe4:hString;
    naturecmd,qualifpiece,modegest,extra:hString;
begin
 //initialisation
 result:='';
 axe1:='';
 axe2:='';
 axe3:='';
 axe4:='';
 valaxe1:='';
 valaxe2:='';
 valaxe3:='';
 valaxe4:='';
 naturecmd:='';
 qualifpiece:='';
 modegest:='';
 extra:='';
 CodeJoin:='';
 //récupère les restrictions de la session donnée
 Q:=MOPenSql('SELECT QBS_CODEAXE1,QBS_CODEAXE2,QBS_CODEAXE3,QBS_CODEAXE4,'+
             'QBS_VALEURAXE1,QBS_VALEURAXE2,QBS_VALEURAXE3,QBS_VALEURAXE4,'+
             'QBS_NATURECMD,QBS_NATURE2,QBS_CODESTRUCT,QBS_VALRESTRICTION '+
             'FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (ChercheCodeRestrictionSession).',true);
 if not Q.eof
  then
   begin
    axe1:=Q.fields[0].asString;
    axe2:=Q.fields[1].asString;
    axe3:=Q.fields[2].asString;
    axe4:=Q.fields[3].asString;
    valaxe1:=Q.fields[4].asString;
    valaxe2:=Q.fields[5].asString;
    valaxe3:=Q.fields[6].asString;
    valaxe4:=Q.fields[7].asString;
    naturecmd:=Q.fields[8].asString;
    qualifpiece:=Q.fields[9].asString;
    structure:=Q.fields[10].asString;
    valRestriction:=Q.fields[11].asString;
   end;
 ferme(Q);

 if BPOkOrli
  then
   //-----------------> ORLI
   begin
    ValRestrictionValaxe1234(valRestriction,valaxe1,valaxe2,valaxe3,valaxe4);
    code:=RechercheCodeRestrictionAutre(codeAs,codeChamp,prefixe,structure,natureCmd,
                                        axe1,axe2,axe3,axe4,
                                        valaxe1,valaxe2,valaxe3,valaxe4);
   end
   //ORLI <-----------------
  else code:=RechercheCodeRestrictionPGI('',natureCmd,qualifpiece,axe1,axe2,axe3,axe4,
                                         valaxe1,valaxe2,valaxe3,valaxe4,CodeJoin);

 result:=code;
end;

//donne les différents codes axes pour la session donnée
procedure ChercheCodeAxeSession(codeSession:hString;
          var TabCodeAxe:array of hString);
var Q:TQuery;
    i:integer;
begin
  Q:=MOPenSql('SELECT '+
             'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4 '+
             ',QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,'+
             'QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10 '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'" ',
             'BPFctSession (ChercheCodeAxeSession).',true);
  if not Q.eof then
    for i:=0 to 9 do
     TabCodeAxe[i+1]:=Q.fields[i].asString;
  ferme(Q);
end;

//cherche dans la structure de la table pivot
//où se trouvent les codes choisis dans le profil
procedure ChercheDixNumValAxe(const structure:hString;
          TabCodeAxe:array of hString;
          var TabNumValAxe:array of hString);
var Q:TQuery;
    i,j:integer;
    CodeAxe: hString;
begin
  if (structure='#PGIENT') or (structure='#PGI') then exit;
  Q:=MOpenSql('SELECT QBC_CODEAXES1,QBC_CODEAXES2,'+
             'QBC_CODEAXES3,QBC_CODEAXES4,QBC_CODEAXES5,QBC_CODEAXES6,'+
             'QBC_CODEAXES7,QBC_CODEAXES8,QBC_CODEAXES9,QBC_CODEAXES10,'+
             'QBC_CODEAXES11,QBC_CODEAXES12,QBC_CODEAXES13,'+
             'QBC_CODEAXES14,QBC_CODEAXES15,QBC_CODEAXES16,'+
             'QBC_CODEAXES17,QBC_CODEAXES18,QBC_CODEAXES19,'+
             'QBC_CODEAXES20,QBC_CODEAXES21,QBC_CODEAXES22,'+
             'QBC_CODEAXES23,QBC_CODEAXES24,QBC_CODEAXES25,'+
             'QBC_CODEAXES26,QBC_CODEAXES27,QBC_CODEAXES28,'+
             'QBC_CODEAXES29,QBC_CODEAXES30 '+
             ' FROM QBPSTRUCTURE WHERE QBC_CODESTRUCT="'+structure+'"',
             'BPFctSession (ChercheDixNumValAxe).',true);
  if not Q.eof then
  begin
    for j:= 1 to 10 do
    begin
      CodeAxe:= TabCodeAxe[j];
      if CodeAxe='' then continue;
      for i:=0 to 29 do
        if CodeAxe=Q.Fields[i].asString then TabNumValAxe[j]:=IntToStr(i+1);
    end;
  end;
  ferme(Q);
end;

//retourne le code structure de la session donnée
function ChercheSessionStructure(const codeSession:hString):hString;
var Q:TQuery;
begin
 result:='';
 Q:=MOpenSql('SELECT QBS_CODESTRUCT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'"','',true);
 if not Q.eof
  then result:=Q.fields[0].asString;
 ferme(Q);
end;

function existCodeAxe(const TabCodeAxe:array of hString; const CodeAxe: hString): boolean;
var i: integer;
begin
  Result:= false;
  for i:=low(TabCodeAxe) to high(TabCodeAxe) do
   if TabCodeAxe[i]=CodeAxe then
   begin
     Result:= true;
     exit;
   end;
end;

procedure ChercheTabCodeAxeTabNumValAxe(const codeSession:hString;
          var TabCodeAxe,TabNumValAxe:array of hString);
var structure:hString;
begin
 structure:=ChercheSessionStructure(codeSession);
 ChercheCodeAxeSession(codeSession,TabCodeAxe);
 ChercheDixNumValAxe(structure,TabCodeAxe,TabNumValAxe);
end;


//donne les différents codes axes
//suivant le code session donné
procedure ChercheCodeAxeSessionTStL(codeSession:hString;
          TStCodeAxe:TStringList);
var Q:TQuery;
    i:integer;
begin
  Q:=MOPenSql('SELECT '+
             'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4 '+
             ',QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,'+
             'QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10 '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'" ',
             'BPFctSession (ChercheCodeAxeSessionTStL).',true);
  if not Q.eof then
    for i:=0 to 9 do
     TStCodeAxe.Add(Q.fields[i].asString);
  ferme(Q);
end;


//procedure qui donne
//pour le code session donné
function ChercheSessionObjectif(const session:hString;
          var SessionObjectif:hString):boolean;
var Q:TQuery;
begin
 result:=false;
 SessionObjectif:='';
 Q:=MOPenSql('SELECT QBS_SESIONOBJECTIF '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (ChercheSessionObjectif).',true);
 if not Q.eof then
 begin
    SessionObjectif:=Q.fields[0].asString;
    if SessionObjectif<>''
     then result:=true;
 end;
 ferme(Q);
end;

//période et date

procedure DonneDateDebEtFinSaison(const Saison:hString;
          var dateDeb,dateFin:TDateTime);
var Q:TQuery;
begin
 dateDeb:=0;
 dateFin:=0;
 Q:=MOpenSql('SELECT QSA_SAISONDEBUT,QSA_SAISONFIN FROM QSAISON '+
             'WHERE '+whereCtx(['qsaison'])+' AND QSA_SAISON="'+Saison+'"',
             'DonneDateDebEtFinSaison (BPFctSession).',true);
 if not Q.eof
  then
   begin
    dateDeb:=Q.fields[0].asDateTime;
    dateFin:=Q.fields[1].asDateTime;
   end;
 ferme(Q);
end;


//donne les dates de début et de fin des périodes courante et de référence
//de la session donnée
procedure ChercheDateDDateFPeriode(const session:hString;
          var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime);
var Q:TQuery;
    saisonCmdC,saisonCmdRef:hString;
begin
 DateDebC:=0;
 DateFinC:=0;
 DateDebRef:=0;
 DateFinRef:=0;
 Q:=MOPenSql('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,'+
             'QBS_DATEFINREF,QBS_SAISONCMDC,QBS_SAISONCMDREF '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (ChercheDateDDateFPeriode).',true);
 if not Q.eof
  then
   begin
    DateDebC:=Q.fields[0].asDateTime;
    DateFinC:=Q.fields[1].asDateTime;
    DateDebRef:=Q.fields[2].asDateTime;
    DateFinRef:=Q.fields[3].asDateTime;
    saisonCmdC:=Q.fields[4].asString;
    saisonCmdRef:=Q.fields[5].asString;
   end;
 ferme(Q);
 if DateDebC=0
  then
   begin
    DonneDateDebEtFinSaison(saisonCmdC,DateDebC,DateFinC);
    DonneDateDebEtFinSaison(saisonCmdRef,DateDebRef,DateFinRef);
   end;
end;

//procedure qui donne la saison de commande et la saison de cmd de référence
//pour le code session donné
procedure ChercheSaisonCmd(const session:hString;
          var SaiCmd,SaiCmdRef:hString);
var Q:TQuery;
begin
 SaiCmd:='';
 SaiCmdRef:='';
 Q:=MOPenSql('SELECT QBS_SAISONCMDC,QBS_SAISONCMDREF '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (ChercheSaisonCmd).',true);
 if not Q.eof
  then
   begin
    SaiCmd:=Q.fields[0].asString;
    SaiCmdRef:=Q.fields[1].asString;
   end;
 ferme(Q);
end;

//****************************************************************************
//ok session ...

//retourne vrai si la session est un objectif
//faux si c'est une prevision
function OkSessionObjectif(const codesession:hString):boolean;
begin
 result:=true;
 if ExisteSql('SELECT QBS_TYPENATURE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codesession+
             '" AND QBS_TYPENATURE="2" ')
  then result:=false;
end;

//retourne vrai si session a été modifiée
//càd si coche okmodifsession=true
function  OkSessionModif(const codeSession:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_OKMODIFSESSION FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+
             '" AND (QBS_OKMODIFSESSION="X" OR  QBS_OKMODIFSESSION is null)')
  then result:=true;
end;

//retourne vrai si session pour permanent
//faux si saisonnier
function  OkSessionPermanent(const codeSession:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_MODEGESTPREV FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+
             '" AND QBS_MODEGESTPREV="2" ')
  then result:=true;
end;

function  OkSessionModifVue(const codeSession:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_OKMODIFVUE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+
             '" AND (QBS_OKMODIFVUE="X" OR  QBS_OKMODIFVUE is null)')
  then result:=true;
end;


//****************************************************************************

//donne la valeur à affichée dans arbre
//cas PGI
function DonneValeurAffiche(const codeSession:hString):integer;
var Q:TQuery;
    RI:integer;
    RS:hString;
begin
  RI:=1;
  Q:=MOpenSql('SELECT QBS_VUEARBRE FROM QBPSESSIONBP '+
              'WHERE QBS_CODESESSION="'+codesession+'"','BPFctSession (DonneValeurAffiche).',true);
  if not Q.eof then RS:=Q.fields[0].asString;
  ferme(Q);
  
  Case ContextBP of
    0,1 : begin //Mode-GC
          if RS='PTC' then RI:=1;
          if RS='QTE' then RI:=2;
          if RS='PHT' then RI:=3;
          if RS='UHT' then RI:=4;
          if RS='UTC' then RI:=5;
          if RS='PAH' then RI:=6;
          if RS='MAR' then RI:=7;
        end;
    2 : begin //Compta
          if RS='DC1' then RI:=1;
          if RS='CD1' then RI:=3;
        end;
    3 : begin //Paie
          RI:=0;
        end
  end;

  result:=RI;
end;

//****************************************************************************
//session etat

//retourne vrai si une session est eclatee par delai ou init par delai
function SessionDelai(const session:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_SESSIONECLAT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+
             '" AND ((QBS_SESSIONECLAT="X") OR (QBS_INITDELAI="X"))')
  then result:=true;
end;

//retourne vrai si la session donnée a déjà été validée
//càd si sa date de validation <>0
function SessionValide(const session:hString):boolean;
var Q:TQuery;
begin
 result:=false;
 Q:=MOPenSql('SELECT QBS_DATEVALIDATION FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (SessionValide).',true);
 if not Q.eof then result:=(Q.fields[0].asDateTime)>10;
 ferme(Q);
end;

//Propose une mise à jour vers la nouvelle version si nécessaire
//permet de générer le détail de l'arbre en gardant les objectifs initiaux
procedure SessionMAJ(const session:hString);
var Q:TQuery;
rep:integer;
InitOK:hString;
DateDebC,DateFinC,DateDebRef,DateFinRef:hString;
begin
 Q:=MOPenSql('SELECT QBS_INITDELAI,QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'"',
             'BPFctSession (SessionValide).',true);
 if not Q.eof then
 begin
   InitOK:=Q.fields[0].asString;
   DateDebC:=Q.fields[1].asString;
   DateFinC:=Q.fields[2].asString;
   DateDebRef:=Q.fields[3].asString;
   DateFinRef:=Q.fields[4].asString;
 end;
 ferme(Q);

 rep:=MrNo;            //Pour le cas où une session ne renvoit aucun résultat
 if ((InitOK='X') AND (ExisteSql('SELECT ##TOP 1## QBR_CODESESSION FROM QBPARBRE WHERE QBR_CODESESSION="'+session+'"'))) then
 begin
   if not ExisteSql('SELECT ##TOP 1## QBH_CODESESSION FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+session+'"')
   then rep:=HShowmessage('1;Session Objectif;Cette session nécessite une mise à jour.'+#13#10+' Voulez-vous effectuer le traitement ?;Q;YN;N;N', '', '');
   if rep = mrYes
   then FuncSynBPLanceCalculInitObjectif([session,'X','','',DateDebC,DateFinC,DateDebRef,DateFinRef],0);
 end;
end;

//retourne vrai si une session a été calculée par délai
//et le mode d'initialisation mois/semaine/quinzaine
function SessionCalculParDelai(const session:hString;
         var BPInitialise:hString):boolean;
var Q:TQuery;
    BP:hString;
begin
 result:=false;
 BPInitialise:='0';
 Q:=MOPenSql('SELECT QBS_BPINITIALISE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'" AND QBS_METHODE="1"',
             'BPFctSession (SessionCalculParDelai).',true);
 if not Q.eof
  then
   begin
    BP:=Q.fields[0].asString;
    if (BP<>'0') and (BP<>'')
     then
      begin
       result:=true;
       BPInitialise:=BP;
      end;
   end;
 ferme(Q);
end;

//retourne vrai si la session de prevision
//à la méthode definie par taille
function SessionCalculParTaille(const codeSession:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_BPINITIALISE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'" AND QBS_METHODETAILLE="1"')
 then result:=true;
end;

//retourne vrai si une session est initialisée prevision
function  SessionInitPrev(const session:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_SESSIONINIT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+
             '" AND QBS_INITPREVISION="X"')
  then result:=true;
end;

//retourne vrai si une session est initialisée coefficient
function  SessionInitCoeff(const session:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_SESSIONINIT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+
             '" AND QBS_INITCOEFF="X"')
  then result:=true;
end;

//retourne vrai si une session est eclatee par delai
function SessionEclateeParDelai(const session:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_SESSIONECLAT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+
             '" AND QBS_SESSIONECLAT="X"')
  then result:=true;
end;

//retourne vrai si une session est eclatee par delai
function SessionEclateeParTaille(const session:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBS_SESSIONECLAT FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+
             '" AND QBS_SESSIONECLATAI="X"')
  then result:=true;
end;


//retourne vrai si une session est definie pour le maillage jour
//cad si le premier axe de la session est etablissement
//et si la session est initialisée en semaine ou en mois 445
function SessionMaillageJour(const codeSession:hString):boolean;
var Q:TQuery;
    codeAxe,BPInitialise,SessionEclat,SessionDelai:hString;
begin
 result:=false;
 BPInitialise:='';
 codeAxe:='';
 SessionEclat:='-';
 SessionDelai:='-';
 Q:=MOPenSql('SELECT QBS_CODEAXES1,QBS_BPINITIALISE,QBS_SESSIONECLAT,QBS_INITDELAI'+
             ' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+
             '" ','BPFctSession (SessionMaillageJour).',true);
 if not Q.eof
  then
   begin
    codeAxe:=Q.fields[0].asString;
    BPInitialise:=Q.fields[1].asString;
    SessionEclat:=Q.fields[2].asString;
    SessionDelai:=Q.fields[3].asString;
   end;
 ferme(Q);
 if (SessionDelai='X') and ((BPInitialise='2') or (BPInitialise='4')) and (codeAxe='003')
  then result:=true;
 if (SessionEclat='X') and (codeAxe='003')
  then result:=true;
end;



//****************************************************************************

//mets à jour le champ okmodif
//avec la valeur Modif (X : true, - : false)
procedure ModifSession(const codeSession,Modif:hString);
begin
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_OKMODIFSESSION="'+Modif+'" '+
            ' WHERE QBS_CODESESSION="'+codeSession+
            '" ','BPFctSession (ModifSession).');
end;

//mets à jour le champ okmodif vue
//avec la valeur Modif (X : true, - : false)
procedure ModifSessionVue(const codeSession,Modif,ModifChpSession:hString);
begin
 if codeSession=''
  then
   if ModifChpSession='X'
    then MExecuteSql('UPDATE QBPSESSIONBP SET QBS_OKMODIFVUE="'+Modif+'" ',
                   'BPFctSession (ModifSessionVue).')
  else
   begin
    if ModifChpSession='X'
     then MExecuteSql('UPDATE QBPSESSIONBP SET QBS_OKMODIFVUE="'+Modif+'" '+
                ' WHERE QBS_CODESESSION="'+codeSession+
                '" ','BPFctSession (ModifSessionVue).');
    //mise à jour de ModifSessionVue pour les sessions de prev
    //qui sont definies à partir de la session d'objectif donnée 
    MExecuteSql('UPDATE QBPSESSIONBP SET QBS_OKMODIFVUE="'+Modif+'" '+
                ' WHERE QBS_SESIONOBJECTIF="'+codeSession+
                '" ','BPFctSession (ModifSessionVue).');
   end;
end;

//****************************************************************************

function ChercheClientArbreSession(const codeSession:hString;
         nivMax:integer):hString;
var Q:TQuery;
    i:integer;
    code:hString;
begin
 result:='';
 Q:=MOpenSql('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4'+
              ',QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,'+
              'QBS_CODEAXES9,QBS_CODEAXES10 FROM QBPSESSIONBP '+
              'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPFctSession (ChercheAxeStructure).',true);
 if not Q.eof
  then
   begin
    for i:=0 to 9 do
     begin
      code:=Q.fields[i].asString;
      if pos('CLIENT',uppercase(code))<>0
       then
        if nivmax=i+1
         then result:='QBR_VALEURAXE '
         else result:='QBR_VALAXENIV'+IntToStr(i+1)+' ';
     end;
   end;
 ferme(Q);
end;


procedure DuplicationArbreBP(const codeSession,codeSessionN:hString);
begin
 MExecuteSql('INSERT INTO QBPARBRE (QBR_CODEAXE,QBR_CODESESSION,QBR_COEFFCALCUL,'+
            'QBR_COEFFGENE,QBR_COEFFRETENU,QBR_COMMENTAIREBP,QBR_DATEDELAI,'+
            'QBR_DEVISE,QBR_EVOLPRCT1,QBR_EVOLPRCT2,QBR_EVOLPRCT3,'+
            'QBR_EVOLPRCT4,QBR_EVOLPRCT5,QBR_EVOLPRCT6,QBR_EVOLQTE,'+
            'QBR_EVOLQTEPRCT,QBR_EVOLVAL1,QBR_EVOLVAL2,QBR_EVOLVAL3,'+
            'QBR_EVOLVAL4,QBR_EVOLVAL5,QBR_EVOLVAL6,QBR_HISTO,'+
            'QBR_LIBVALAXE,QBR_NBCLTNOUVEAU,QBR_NBCLTPERDU,QBR_NBCLTREF,'+
            'QBR_NBCLTRESTAVOIR,QBR_NBCLTVU,QBR_NEWPROSPECT,QBR_NIVEAU,'+
            'QBR_NOUVEAU,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_OP1,QBR_OP2,'+
            'QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6,QBR_OPPRCT1,QBR_OPPRCT2,'+
            'QBR_OPPRCT3,QBR_OPPRCT4,QBR_OPPRCT5,QBR_OPPRCT6,QBR_PERDU,'+
            'QBR_PREVU,QBR_PROSPECT,QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,'+
            'QBR_QTEREFPRCT,QBR_REALISE,QBR_REF1,QBR_REF2,QBR_REF3,'+
            'QBR_REF4,QBR_REF5,QBR_REF6,QBR_REFPRCT1,QBR_REFPRCT2,'+
            'QBR_REFPRCT3,QBR_REFPRCT4,QBR_REFPRCT5,QBR_REFPRCT6,'+
            'QBR_RESTEAVISTER,QBR_SAISI1,QBR_SAISI2,QBR_SAISI3,QBR_SAISI4,'+
            'QBR_SAISI5,QBR_SAISI6,QBR_SAISIQTE,QBR_TOTALCOURANT,QBR_TOTALREF,'+
            'QBR_VALAXENIV1,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,'+
            'QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,QBR_VALAXENIV2,'+
            'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
            'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALBLOQUE,'+
            'QBR_VALBLOQUETMP,QBR_VALBLOQUETMP1,QBR_VALEURAXE,QBR_VUCOURANT,'+
            'QBR_VUREF,QBR_XVISIT) '+
            'SELECT QBR_CODEAXE,"'+CodeSessionN+'",QBR_COEFFCALCUL,'+
            'QBR_COEFFGENE,QBR_COEFFRETENU,QBR_COMMENTAIREBP,QBR_DATEDELAI,'+
            'QBR_DEVISE,QBR_EVOLPRCT1,QBR_EVOLPRCT2,QBR_EVOLPRCT3,'+
            'QBR_EVOLPRCT4,QBR_EVOLPRCT5,QBR_EVOLPRCT6,QBR_EVOLQTE,'+
            'QBR_EVOLQTEPRCT,QBR_EVOLVAL1,QBR_EVOLVAL2,QBR_EVOLVAL3,'+
            'QBR_EVOLVAL4,QBR_EVOLVAL5,QBR_EVOLVAL6,QBR_HISTO,'+
            'QBR_LIBVALAXE,QBR_NBCLTNOUVEAU,QBR_NBCLTPERDU,QBR_NBCLTREF,'+
            'QBR_NBCLTRESTAVOIR,QBR_NBCLTVU,QBR_NEWPROSPECT,QBR_NIVEAU,'+
            'QBR_NOUVEAU,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_OP1,QBR_OP2,'+
            'QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6,QBR_OPPRCT1,QBR_OPPRCT2,'+
            'QBR_OPPRCT3,QBR_OPPRCT4,QBR_OPPRCT5,QBR_OPPRCT6,QBR_PERDU,'+
            'QBR_PREVU,QBR_PROSPECT,QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,'+
            'QBR_QTEREFPRCT,QBR_REALISE,QBR_REF1,QBR_REF2,QBR_REF3,'+
            'QBR_REF4,QBR_REF5,QBR_REF6,QBR_REFPRCT1,QBR_REFPRCT2,'+
            'QBR_REFPRCT3,QBR_REFPRCT4,QBR_REFPRCT5,QBR_REFPRCT6,'+
            'QBR_RESTEAVISTER,QBR_SAISI1,QBR_SAISI2,QBR_SAISI3,QBR_SAISI4,'+
            'QBR_SAISI5,QBR_SAISI6,QBR_SAISIQTE,QBR_TOTALCOURANT,QBR_TOTALREF,'+
            'QBR_VALAXENIV1,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,'+
            'QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,QBR_VALAXENIV2,'+
            'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
            'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALBLOQUE,'+
            'QBR_VALBLOQUETMP,QBR_VALBLOQUETMP1,QBR_VALEURAXE,QBR_VUCOURANT,'+
            'QBR_VUREF,QBR_XVISIT FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"',
            'BPFctSession (DuplicationSessionBP).');
MExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
            'QBH_DATEDELAI,QBH_QTEREF,QBH_QTEREFPRCT,QBH_QTEC,QBH_QTECPRCT,'+
            'QBH_REF1,QBH_REFPRCT1,QBH_OP1,QBH_OPPRCT1,'+
            'QBH_REF2,QBH_REFPRCT2,QBH_OP2,QBH_OPPRCT2,'+
            'QBH_REF3,QBH_REFPRCT3,QBH_OP3,QBH_OPPRCT3,'+
            'QBH_REF4,QBH_REFPRCT4,QBH_OP4,QBH_OPPRCT4,'+
            'QBH_REF5,QBH_REFPRCT5,QBH_OP5,QBH_OPPRCT5,'+
            'QBH_REF6,QBH_REFPRCT6,QBH_OP6,QBH_OPPRCT6,'+
            'QBH_DATEPIECE,QBH_NUMNOEUDREF,QBH_PRCTVARIATION1,QBH_PRCTVARIATION2,'+
            'QBH_PRCTVARIATION3,QBH_PRCTVARIATION4,QBH_PRCTVARIATION5,QBH_PRCTVARIATION6,'+
            'QBH_PRCTVARIATIONQ) '+
            'SELECT "'+CodeSessionN+'",QBH_NUMNOEUD,QBH_NIVEAU,'+
            'QBH_DATEDELAI,QBH_QTEREF,QBH_QTEREFPRCT,QBH_QTEC,QBH_QTECPRCT,'+
            'QBH_REF1,QBH_REFPRCT1,QBH_OP1,QBH_OPPRCT1,'+
            'QBH_REF2,QBH_REFPRCT2,QBH_OP2,QBH_OPPRCT2,'+
            'QBH_REF3,QBH_REFPRCT3,QBH_OP3,QBH_OPPRCT3,'+
            'QBH_REF4,QBH_REFPRCT4,QBH_OP4,QBH_OPPRCT4,'+
            'QBH_REF5,QBH_REFPRCT5,QBH_OP5,QBH_OPPRCT5,'+
            'QBH_REF6,QBH_REFPRCT6,QBH_OP6,QBH_OPPRCT6,'+
            'QBH_DATEPIECE,QBH_NUMNOEUDREF,QBH_PRCTVARIATION1,QBH_PRCTVARIATION2,'+
            'QBH_PRCTVARIATION3,QBH_PRCTVARIATION4,QBH_PRCTVARIATION5,QBH_PRCTVARIATION6,'+
            'QBH_PRCTVARIATIONQ FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"',
            'BPFctSession (DuplicationSessionBP).');
end;

procedure DuplicationArbrePartielleBP(const codeSession,codeSessionN,Niv:hString);
begin
 MExecuteSql('INSERT INTO QBPARBRE (QBR_CODEAXE,QBR_CODESESSION,QBR_COEFFCALCUL,'+
            'QBR_COEFFGENE,QBR_COEFFRETENU,QBR_COMMENTAIREBP,QBR_DATEDELAI,'+
            'QBR_DEVISE,QBR_EVOLPRCT1,QBR_EVOLPRCT2,QBR_EVOLPRCT3,'+
            'QBR_EVOLPRCT4,QBR_EVOLPRCT5,QBR_EVOLPRCT6,QBR_EVOLQTE,'+
            'QBR_EVOLQTEPRCT,QBR_EVOLVAL1,QBR_EVOLVAL2,QBR_EVOLVAL3,'+
            'QBR_EVOLVAL4,QBR_EVOLVAL5,QBR_EVOLVAL6,QBR_HISTO,'+
            'QBR_LIBVALAXE,QBR_NBCLTNOUVEAU,QBR_NBCLTPERDU,QBR_NBCLTREF,'+
            'QBR_NBCLTRESTAVOIR,QBR_NBCLTVU,QBR_NEWPROSPECT,QBR_NIVEAU,'+
            'QBR_NOUVEAU,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_OP1,QBR_OP2,'+
            'QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6,QBR_OPPRCT1,QBR_OPPRCT2,'+
            'QBR_OPPRCT3,QBR_OPPRCT4,QBR_OPPRCT5,QBR_OPPRCT6,QBR_PERDU,'+
            'QBR_PREVU,QBR_PROSPECT,QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,'+
            'QBR_QTEREFPRCT,QBR_REALISE,QBR_REF1,QBR_REF2,QBR_REF3,'+
            'QBR_REF4,QBR_REF5,QBR_REF6,QBR_REFPRCT1,QBR_REFPRCT2,'+
            'QBR_REFPRCT3,QBR_REFPRCT4,QBR_REFPRCT5,QBR_REFPRCT6,'+
            'QBR_RESTEAVISTER,QBR_SAISI1,QBR_SAISI2,QBR_SAISI3,QBR_SAISI4,'+
            'QBR_SAISI5,QBR_SAISI6,QBR_SAISIQTE,QBR_TOTALCOURANT,QBR_TOTALREF,'+
            'QBR_VALAXENIV1,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,'+
            'QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,QBR_VALAXENIV2,'+
            'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
            'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALBLOQUE,'+
            'QBR_VALBLOQUETMP,QBR_VALBLOQUETMP1,QBR_VALEURAXE,QBR_VUCOURANT,'+
            'QBR_VUREF,QBR_XVISIT) '+
            'SELECT QBR_CODEAXE,"'+CodeSessionN+'",QBR_COEFFCALCUL,'+
            'QBR_COEFFGENE,QBR_COEFFRETENU,QBR_COMMENTAIREBP,QBR_DATEDELAI,'+
            'QBR_DEVISE,QBR_EVOLPRCT1,QBR_EVOLPRCT2,QBR_EVOLPRCT3,'+
            'QBR_EVOLPRCT4,QBR_EVOLPRCT5,QBR_EVOLPRCT6,QBR_EVOLQTE,'+
            'QBR_EVOLQTEPRCT,QBR_EVOLVAL1,QBR_EVOLVAL2,QBR_EVOLVAL3,'+
            'QBR_EVOLVAL4,QBR_EVOLVAL5,QBR_EVOLVAL6,QBR_HISTO,'+
            'QBR_LIBVALAXE,QBR_NBCLTNOUVEAU,QBR_NBCLTPERDU,QBR_NBCLTREF,'+
            'QBR_NBCLTRESTAVOIR,QBR_NBCLTVU,QBR_NEWPROSPECT,QBR_NIVEAU,'+
            'QBR_NOUVEAU,QBR_NUMNOEUD,QBR_NUMNOEUDPERE,QBR_OP1,QBR_OP2,'+
            'QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6,QBR_OPPRCT1,QBR_OPPRCT2,'+
            'QBR_OPPRCT3,QBR_OPPRCT4,QBR_OPPRCT5,QBR_OPPRCT6,QBR_PERDU,'+
            'QBR_PREVU,QBR_PROSPECT,QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,'+
            'QBR_QTEREFPRCT,QBR_REALISE,QBR_REF1,QBR_REF2,QBR_REF3,'+
            'QBR_REF4,QBR_REF5,QBR_REF6,QBR_REFPRCT1,QBR_REFPRCT2,'+
            'QBR_REFPRCT3,QBR_REFPRCT4,QBR_REFPRCT5,QBR_REFPRCT6,'+
            'QBR_RESTEAVISTER,QBR_SAISI1,QBR_SAISI2,QBR_SAISI3,QBR_SAISI4,'+
            'QBR_SAISI5,QBR_SAISI6,QBR_SAISIQTE,QBR_TOTALCOURANT,QBR_TOTALREF,'+
            'QBR_VALAXENIV1,QBR_VALAXENIV10,QBR_VALAXENIV11,QBR_VALAXENIV12,'+
            'QBR_VALAXENIV13,QBR_VALAXENIV14,QBR_VALAXENIV15,QBR_VALAXENIV2,'+
            'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
            'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALBLOQUE,'+
            'QBR_VALBLOQUETMP,QBR_VALBLOQUETMP1,QBR_VALEURAXE,QBR_VUCOURANT,'+
            'QBR_VUREF,QBR_XVISIT FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NIVEAU<="'+STRFPOINT(VALEURI(niv)+1)+'"',
            'BPFctSession (DuplicationSessionBP).');
MExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
            'QBH_DATEDELAI,QBH_QTEREF,QBH_QTEREFPRCT,QBH_QTEC,QBH_QTECPRCT,'+
            'QBH_REF1,QBH_REFPRCT1,QBH_OP1,QBH_OPPRCT1,'+
            'QBH_REF2,QBH_REFPRCT2,QBH_OP2,QBH_OPPRCT2,'+
            'QBH_REF3,QBH_REFPRCT3,QBH_OP3,QBH_OPPRCT3,'+
            'QBH_REF4,QBH_REFPRCT4,QBH_OP4,QBH_OPPRCT4,'+
            'QBH_REF5,QBH_REFPRCT5,QBH_OP5,QBH_OPPRCT5,'+
            'QBH_REF6,QBH_REFPRCT6,QBH_OP6,QBH_OPPRCT6,'+
            'QBH_DATEPIECE,QBH_NUMNOEUDREF,QBH_PRCTVARIATION1,QBH_PRCTVARIATION2,'+
            'QBH_PRCTVARIATION3,QBH_PRCTVARIATION4,QBH_PRCTVARIATION5,QBH_PRCTVARIATION6,'+
            'QBH_PRCTVARIATIONQ) '+
            'SELECT "'+CodeSessionN+'",QBH_NUMNOEUD,QBH_NIVEAU,'+
            'QBH_DATEDELAI,QBH_QTEREF,QBH_QTEREFPRCT,QBH_QTEC,QBH_QTECPRCT,'+
            'QBH_REF1,QBH_REFPRCT1,QBH_OP1,QBH_OPPRCT1,'+
            'QBH_REF2,QBH_REFPRCT2,QBH_OP2,QBH_OPPRCT2,'+
            'QBH_REF3,QBH_REFPRCT3,QBH_OP3,QBH_OPPRCT3,'+
            'QBH_REF4,QBH_REFPRCT4,QBH_OP4,QBH_OPPRCT4,'+
            'QBH_REF5,QBH_REFPRCT5,QBH_OP5,QBH_OPPRCT5,'+
            'QBH_REF6,QBH_REFPRCT6,QBH_OP6,QBH_OPPRCT6,'+
            'QBH_DATEPIECE,QBH_NUMNOEUDREF,QBH_PRCTVARIATION1,QBH_PRCTVARIATION2,'+
            'QBH_PRCTVARIATION3,QBH_PRCTVARIATION4,QBH_PRCTVARIATION5,QBH_PRCTVARIATION6,'+
            'QBH_PRCTVARIATIONQ FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+
            '" AND QBH_NIVEAU<="'+STRFPOINT(VALEURI(niv)+1)+'"',
            'BPFctSession (DuplicationSessionBP).');

end;

//****************************************************************************
//methode

//retourne la methode delai de la session
function  DonneMethodeSession(const codeSession:hString):hString;
var Q: TQuery;
begin
  Result:='';
  Q:= openSql('SELECT QBS_METHODE FROM QBPSESSIONBP WHERE '+
              'QBS_CODESESSION="'+codeSession+'"',true);
  if not(Q.eof)
   then Result:=Q.fields[0].asString;
  ferme(Q);
end;

//****************************************************************************

function ChercheNumSession(const codeSession:hString):hString;
var Q:TQuery;
begin
 result:='0';
 Q:=MOpenSql('SELECT QBS_NUMSESSION FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPFctSession (ChercheNumSession).',true);
 if not Q.eof
  then result:=Q.fields[0].asString;
 ferme(Q);
end;

//****************************************************************************

//retourne faux si le code session est incorrect
function ControleSession(session: hString): boolean;
begin
 result:=true;
 if session=''
  then result:=false
  else
   begin
    if not ExisteSQL('SELECT QBS_CODESESSION FROM QBPSESSIONBP '+
                     'WHERE QBS_CODESESSION="'+session+'"')
     then result:=false;
   end;
end;

//****************************************************************************

function IncrementeCodeSession(const codeSession:hString):hString;
var S,numS,codeS,codeX:hString;
    numI,code,i,ii,L:integer;
    Q:TQuery;
begin
 result:=codeSession;
 //cherche le dernier _ dans le code session
 S:=codeSession;
 ii:=0;
 while pos('-',S)<>0 do
  begin
   i:=pos('-',S);
   ii:=ii+i;
   L:=length(S);
   S:=copy(S,i+1,L-i);
  end;
 if ii<>0
  then
   begin
    numS:=copy(codeSession,ii+1,3);
    codeS:=copy(codeSession,1,ii-1);
    val(numS,numI,code);
    if code=0
     then
      begin
       Q:=MOpenSql('SELECT MAX(QBS_CODESESSION) FROM QBPSESSIONBP '+
                   'WHERE QBS_CODESESSION LIKE "'+codeS+'-%"','BPFctSession (IncrementeCodeSession).',true);
       if not Q.eof
        then
         begin
          numS:=copy(Q.fields[0].asString,ii+1,3);
          codeS:=copy(Q.fields[0].asString,1,ii-1);
          val(numS,numI,code);
          if code=0
           then result:=codeS+'-'+MetZero(IntToStr(numI+1),3);
         end
        else result:=codeS+'-'+MetZero(IntToStr(numI+1),3);
       ferme(Q);
      end
     else result:=codeSession+'-001';
   end
  else
   begin
    codeX:=codeSession+'-001';
    if ExisteSql('SELECT QBS_CODESESSION FROM QBPSESSIONBP '+
                 'WHERE QBS_CODESESSION="'+codeX+'"')
     then
      begin
       Q:=MOpenSql('SELECT MAX(QBS_CODESESSION) FROM QBPSESSIONBP '+
                   'WHERE QBS_CODESESSION LIKE "'+codeSession+'-0%"','BPFctSession (IncrementeCodeSession).',true);
       if not Q.eof
        then
         begin
          S:=Q.fields[0].asString;
          i:=pos('-',S);
          numS:=copy(S,i+1,3);
          codeS:=copy(S,1,i-1);
          val(numS,numI,code);
          result:=codeS+'-'+MetZero(IntToStr(numI+1),3);
         end;
       ferme(Q);
      end
     else result:=codeX;
   end;

end;


procedure ChercheObjectifPourUnJour(CodeChpPrct1,CodeChpPrct2,Etablissement,codeSession:hString;
          DateSemDeb:TDateTime;
          var Prevu:double);
var okPrct:boolean;
    Q:TQuery;
begin
 okPrct:=false;             
 //cherche les % correspondant à la semaine
 if codeChpPrct2<>''
  then
   begin
    Q:=MOPenSql('SELECT '+codeChpPrct2+
             ' FROM QBPDETCALENDREP WHERE QBE_CALENDREP="'+CodeSession+
             '" AND QBE_VALAXENIV1="'+Etablissement+'" AND QBE_DATEDEBSEM="'+USDATETIME(DateSemDeb)+'"',
             'BPFctSession (ChercheObjectifPourUnJour).',true);
    if not Q.eof
     then
      begin
       okPrct:=true;
       Prevu:=(Prevu*Q.fields[0].asFloat)/100;
      end;
    ferme(Q);
  end;

 if not okPrct then Prevu:=0;
 //PGIInfo('Aucun calendrier défini pour la caisse sélectionnée.');
{    if codeChpPrct1<>''
     then
      begin
       Q:=MOPenSql('SELECT '+codeChpPrct1+
                ' FROM QBPCALENDREP WHERE QBN_CALENDREP='+
                '(SELECT ET_CALENDREP FROM ETABLISS WHERE ET_ETABLISSEMENT="'+Etablissement+'") ',
                '',true);
       if not Q.eof
        then
         begin
          Prevu:=(Prevu*Q.fields[0].asFloat)/100;
         end;
       ferme(Q);
      end;
    end;
}
end;

function DonneObjectifsEtablissements(codeSession:hString;
         LesEtablissements:TStringList;DateDebut,DateFin:TDateTime;
         var LibelleValeur,Devise:hstring):double;
var Q:TQuery;
    ValeurAffiche,NumSemDeb,NumSemFin,i,j,nbjour,nbjourX,anDeb,anFin:integer;
    codeChp,codeChpWhere,Etablissement,nbjourS,codeChpPrct1,codeChpPrct2:hString;
    Prevu,PrevuFin,PrevuDeb:double;
    jour,DateMoisDeb,DateMoisFin,DateSemDeb,DateSemFin,DateSemDeb2,DateSemFin2,dateMois:TDateTime;
    OkDelai:boolean;
    nivMax,nbSem,k,l,nbMois:integer;
    BPinitialise,codeWhere,calendrier:hString;
    OkCalendException:boolean;
    TabPrct:array of double;
    sommePrct,prevuM:double;
    An,mois,day,anX,moisX,jourX,anDb,moisDb,jourDb,anFn,moisFn,jourFn,anJ,moisJ,jourJ:word;
begin
 Prevu:=0;
 codeChpPrct1:='';
 codeChpPrct2:='';
 ValeurAffiche:=DonneValeurAffiche(codesession);
 case ValeurAffiche of
  1 : begin
       codeChp:='QBR_OP1';
       LibelleValeur:=TraduireMemoire('CA TTC net');
      end;
  2 : begin
       codeChp:='QBR_QTEC';
       LibelleValeur:=TraduireMemoire('Quantité');
      end;
  3 : begin
       codeChp:='QBR_OP2';
       LibelleValeur:=TraduireMemoire('CA HT net');
      end;
  4 : begin
       codeChp:='QBR_OP3';
       LibelleValeur:=TraduireMemoire('CA TTC avant remise');
      end;
  5 : begin
       codeChp:='QBR_OP4';
       LibelleValeur:=TraduireMemoire('CA HT avant remise');
      end;
  6 : begin
       codeChp:='QBR_OP5';
       LibelleValeur:=TraduireMemoire('Cumul achat en PAMP');
      end;
  7 : begin
       codeChp:='QBR_OP6';
       LibelleValeur:=TraduireMemoire('Marge');
      end;                                                     
  end;


 OkDelai:=SessionCalculParDelai(codeSession,BPinitialise);
 nivMax:=ChercheNivMax(codeSession);

 codeChpWhere:=' AND (';
 for i:=0 to LesEtablissements.Count-1 do
  begin
   if i=0
    then codeChpWhere:=codeChpWhere+' (QBR_VALAXENIV1="'+LesEtablissements[i]+'") '
    else codeChpWhere:=codeChpWhere+' OR (QBR_VALAXENIV1="'+LesEtablissements[i]+'") ';
  end;
 codeChpWhere:=codeChpWhere+') ';


 for i:=0 to LesEtablissements.Count-1 do
  begin
   Etablissement:=LesEtablissements[i];
   //si ce n'est pas une session init par delai
   //on cherche le type de maille (eclatement) dans table arbre
   if not OkDelai
    then
     begin
      if nivMax=1
       then codeWhere:=' QBR_VALEURAXE="'+Etablissement+'" '
       else codeWhere:=' QBR_VALAXENIV1="'+Etablissement+'" ';
      Q:=MOpenSql('SELECT QBR_MAILLE FROM QBPARBRE '+
                  'WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NIVEAU="'+STRFPOINT(nivMax)+
                  '" AND '+codeWhere,'BPFctSession (DonneObjectifsEtablissements).',true);
      if not Q.eof
       then BPInitialise:=Q.fields[0].asString;
      ferme(Q);
     end;

   //cas semaine
   if (BPInitialise='2')
    then
     begin
      //Cherche date debut semaine de date debut
      NumSemDeb:=numSemaine(DateDebut,AnDeb);
      DateSemDeb:=PremierJourSemaine(NumSemDeb,AnDeb);
      //Cherche date debut semaine de date fint
      NumSemFin:= numSemaine(DateFin,AnFin);
      DateSemFin:=PremierJourSemaine(NumSemFin,AnFin);

      //pour le cas journée
      if DateDebut=DateFin
       then
        begin
         //cherche le prevu
         Q:=MOpenSql('SELECT SUM('+codeChp+'),QBR_DEVISE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                   '" '+codeChpWhere+' AND QBR_DATEDELAI>="'+USDATETIME(DateSemDeb)+
                   '" AND QBR_DATEDELAI<="'+USDATETIME(DateSemFin)+
                   '" GROUP BY QBR_DEVISE','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
          then
           begin
            Prevu:=Q.fields[0].asFloat;
            Devise:=Q.fields[1].asString;
           end;
         ferme(Q);

         nbjourS:=IntToStr(round(DateDebut-DateSemDeb)+1);
         codeChpPrct1:='QBN_PRCT'+nbjourS;
         codeChpPrct2:='QBE_PRCTJOUR'+nbjourS;
         ChercheObjectifPourUnJour(CodeChpPrct1,CodeChpPrct2,Etablissement,CodeSession,DateSemDeb,prevu);
        end
       else
        begin
         //pour une periode date
         DateSemDeb2:=DateSemDeb;
         if DateSemDeb<>DateDebut
          then
           begin
            NumSemDeb:= numSemaine(DateDebut+7,AnDeb);
            DateSemDeb2:=PremierJourSemaine(NumSemDeb,AnDeb);
           end;
         DateSemFin2:=DateSemFin;
         if DateSemFin<>DateFin
          then
           begin
            NumSemFin:=numSemaine(DateFin-7,AnDeb);
            DateSemFin2:=PremierJourSemaine(NumSemFin,AnFin);
           end;

         //cherche le prevu
         Q:=MOpenSql('SELECT SUM('+codeChp+'),QBR_DEVISE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                  '" '+codeChpWhere+' AND QBR_DATEDELAI>="'+USDATETIME(DateSemDeb2)+
                  '" AND QBR_DATEDELAI<="'+USDATETIME(DateSemFin2)+
                  '" GROUP BY QBR_DEVISE','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
          then
           begin
            Prevu:=Q.fields[0].asFloat;
            Devise:=Q.fields[1].asString;
           end;
         ferme(Q);
         //pour une periode de date
         //pour avoir le prevu on doit rajouter le prevu des journées qui
         //sont en dehors du semaine complete
         PrevuFin:=0;
         PrevuDeb:=0;
         //cherche le prevu de debut
         Q:=MOpenSql('SELECT SUM('+codeChp+') FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
               '" '+codeChpWhere+' AND QBR_DATEDELAI="'+USDATETIME(DateSemDeb)+
               '" GROUP BY QBR_DEVISE','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
          then PrevuDeb:=Q.fields[0].asFloat;
         ferme(Q);
         //cherche le prevu de fin
         Q:=MOpenSql('SELECT SUM('+codeChp+') FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
               '" '+codeChpWhere+' AND QBR_DATEDELAI="'+USDATETIME(DateSemFin)+
               '" GROUP BY QBR_DEVISE','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
          then PrevuFin:=Q.fields[0].asFloat;
         ferme(Q);

         nbjourX:=round(DateDebut-DateSemDeb)+1;
         nbjour:=round(DateSemDeb2-DateSemDeb);
         for j:=nbjourX to nbJour do
          begin
           if codeChpPrct1=''
            then
             begin
              codeChpPrct1:=' QBN_PRCT'+IntToStr(j);
              codeChpPrct2:=' QBE_PRCTJOUR'+IntToStr(j);
             end
            else
             begin
              codeChpPrct1:=codeChpPrct1+'+QBN_PRCT'+IntToStr(j);
              codeChpPrct2:=codeChpPrct2+'+QBE_PRCTJOUR'+IntToStr(j);
             end;
          end;
         ChercheObjectifPourUnJour(CodeChpPrct1,CodeChpPrct2,Etablissement,CodeSession,DateSemDeb,PrevuDeb);

         codeChpPrct1:='';
         nbjourX:=round(DateFin-DateSemFin)+1;
         for j:=1 to nbJourX do
          begin
           if codeChpPrct1=''
            then
             begin
              codeChpPrct1:=' QBN_PRCT'+IntToStr(j);
              codeChpPrct2:=' QBE_PRCTJOUR'+IntToStr(j);
             end
            else
             begin
              codeChpPrct1:=codeChpPrct1+'+QBN_PRCT'+IntToStr(j);
              codeChpPrct2:=codeChpPrct2+'+QBE_PRCTJOUR'+IntToStr(j);
             end;
          end;
         ChercheObjectifPourUnJour(CodeChpPrct1,CodeChpPrct2,Etablissement,CodeSession,DateSemFin,PrevuFin);
         Prevu:=Prevu+PrevuDeb+PrevuFin;
       end;
    end;

    //cas mois
    if (BPInitialise='4')
     then
      begin
       //pour une periode de date
       DecodeDate(DateDebut,anDb,moisDb,jourDb);
       DecodeDate(DateFin,anFn,moisFn,jourFn);
       if (moisFn>=moisDb)
        then nbMois:=MoisFn-MoisDb+1
        else nbMois:=(12-MoisDb+1)+MoisFn+1;

       for l:=0 to nbMois-1 do
        begin
         dateMois:=PLUSMOIS(DateDebut,l);
         //Cherche date debut mois de date debut
         DateMoisDeb:=DEBUTDEMOIS(dateMois);
         //Cherche date debut semaine de date debut
         NumSemDeb:=numSemaine(DateMoisDeb,AnDeb);
         DateSemDeb:=PremierJourSemaine(NumSemDeb,AnDeb);
         //Cherche date debut mois de date debut
         DateMoisFin:=FINDEMOIS(dateMois);
         //Cherche date debut semaine de date fin
         NumSemFin:= numSemaine(DateMoisFin,AnFin);

         DecodeDate(dateMois,anX,moisX,jourX);

         //cherche le nb de semaine du mois
         if (NumSemFin>=NumSemDeb)
          then nbSem:=NumSemFin - NumSemDeb + 1
          else nbSem:=(52 - NumSemDeb + 1)+NumSemFin+1;

         //longueur du tableau des prcts des jours
         SetLength(TabPrct,(7*nbSem)+1);
{
         //cherche calendrier de l'établissement
         calendrier:='';
         Q:=MOpenSql('SELECT ET_CALENDREP FROM ETABLISS '+
                     'WHERE ET_ETABLISSEMENT="'+Etablissement+'"','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
         then calendrier:=Q.fields[0].asString;
         ferme(Q);
}

         //cherche les prcts des jours des semaines du mois
         for j:=1 to nbSem do
          begin
           //OkCalendException:=false;
           //semaine exception ou semaine type
           Q:=MOPenSql('SELECT QBE_PRCTJOUR1,QBE_PRCTJOUR2,QBE_PRCTJOUR3,QBE_PRCTJOUR4,'+
                     'QBE_PRCTJOUR5,QBE_PRCTJOUR6,QBE_PRCTJOUR7 '+
                     ' FROM QBPDETCALENDREP '+
                     'WHERE QBE_CALENDREP="'+CodeSession+
                     '" AND QBE_VALAXENIV1="'+Etablissement+'" AND QBE_DATEDEBSEM="'+USDATETIME(DateSemDeb+((j-1)*7))+'" ',
                     'BPFctSession (DonneObjectifsEtablissements).',true);
           if not Q.eof
            then
             begin
              //OkCalendException:=true;
              for k:=1 to 7 do
               begin
                TabPrct[k+((j-1)*7)]:=Q.fields[k-1].asFloat;
               end;
             end;
           ferme(Q);

{
           if not OkCalendException
            then
             begin
              //cherche semaine type du calendrier
              Q:=MOpenSql('SELECT QBN_PRCT1,QBN_PRCT2,QBN_PRCT3,QBN_PRCT4,'+
                        'QBN_PRCT5,QBN_PRCT6,QBN_PRCT7 FROM QBPCALENDREP '+
                        'WHERE QBN_CALENDREP="'+Calendrier+'"',
                        'BPFctSession (DonneObjectifsEtablissements).',true);
              if not Q.eof
               then
                begin
                 for k:=1 to 7 do
                  begin
                   TabPrct[k+((j-1)*7)]:=Q.fields[k-1].asFloat;
                  end;
                end;
              ferme(Q);
             end;
}          end;

         //cherche la somme totale des prcts du mois
         sommePrct:=0;
         for k:=1 to nbSem*7 do
          begin
           jour:=PLUSDATE(DateSemDeb,k-1,'J');
           DecodeDate(jour,An,mois,day);
           if mois=MoisX
            then sommePrct:=sommePrct+TabPrct[k];
          end;

         PrevuM:=0;
         //cherche le prevu
         Q:=MOpenSql('SELECT SUM('+codeChp+'),QBR_DEVISE FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                     '" '+codeChpWhere+' AND QBR_DATEDELAI="'+USDATETIME(DateMoisDeb)+
                     '" GROUP BY QBR_DEVISE','BPFctSession (DonneObjectifsEtablissements).',true);
         if not Q.eof
          then
           begin
            PrevuM:=Q.fields[0].asFloat;
            Devise:=Q.fields[1].asString;
           end;
         ferme(Q);

         //dans le cas journée
         if DateDebut=DateFin
          then
           begin
            //cherche la valeur pour la date donnée
            for k:=1 to nbSem*7 do
             begin
              jour:=PLUSDATE(DateSemDeb,k-1,'J');
              if sommePrct<>0
               then
                if jour=DateDebut
               then prevu:=prevuM*(TabPrct[k]/sommePrct);
             end;
           end
          else
           begin
            //pour une periode de date
            for k:=1 to nbSem*7 do
             begin
              jour:=PLUSDATE(DateSemDeb,k-1,'J');
              decodedate(jour,anJ,moisJ,jourJ);
              if moisX=moisJ
               then
                if sommePrct<>0
                 then
                  if DateAppartIntervalle(jour,DateDebut,DateFin)
                   then prevu:=prevu+(prevuM*(TabPrct[k]/sommePrct));
             end;
           end;

        end;


      end;
  end;

 result:=Prevu;
end;

end.
