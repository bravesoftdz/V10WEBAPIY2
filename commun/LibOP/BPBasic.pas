unit BPBasic;

interface

uses  
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$else}
     uTob,
{$ENDIF}
     Classes,Controls,SysUtils,HCtrls,BPFctSession;

function GestionBPColoris:boolean;
function GestionBPFS:boolean;
function GestionBPTaille:boolean;
function GestionBPMagasin:boolean;

function  BPOkOrli:boolean;

function RemplaceCotteEspacePourSql(const S:hString):hString;

procedure DonneCodeChpJoinPGINbNivLib(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,champsGB,joins:hString; const CodeSession: hString);

function DonneLibelleValeurAxe(const CodeAxe,ValeurAxe: hString): hString;
function DonneCodeAxeNiv(const session,Niv: hString): hString;
procedure ChercheCodeSqlChpAxeLib(nivMax:integer;
          TabCodeAxe:array of hString;
          var chpSql,joinSql:hString);
procedure ChercheCodeSqlChpAxeLibAvecPivot(nivMax:integer;const codeAs:hString;
          TabCodeAxe,TabNumValAxe:array of hString;
          var chpSql,joinSql:hString);
procedure DonneCodeChpJoinPGINbNiv(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,joins:hString);
procedure DonneCodeChpJoinPGINbNivP(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs:hString);
procedure ChercheCodeSqlChpAxeLibPGI(nivMax:integer; InitTree:boolean;
          TabCodeAxe:array of hString;
          var chpSql,joinSql:hString);
procedure DonneCodeChpJoinPGINbNivAs(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,joins:hString);
procedure DonneCodeChpJoinPGI(const codeAxe,codeAs:hString;
          var chp,join:hString);

procedure DonneCodeChpUserRestrict(TabCodeAxe:array of hString;
          UserRestrict:hString; niveauI:integer;
          var UserSQLRestrict:hString);

function CalculCoeffQte:boolean;

procedure MiseAJourTabletteValeurAAfficher;

procedure RechercheTabAff(var TabValAff,TabLibValAff:array of hString);

procedure ChangeListeLibelleColonnes(OkObjectif:boolean;Session:string;TabValAff,TabLibValAff:array of hString);

procedure DonneNumTrimestre(DateT:TDateTime;var NumTrimestre,Annee:integer);
procedure DonneNumQuadrimestre(DateT:TDateTime;var NumQuadrimestre,Annee:integer);      
procedure DonneDateDebFinTrimestre(NumT,AnT:integer;var DateDeb,DateFin:TDateTime);
procedure DonneDateDebFinQuadrimestre(NumT,AnT:integer;var DateDeb,DateFin:TDateTime);
function  DonneNbIntervalleMailleDateDebDateFin(const maille:hString;
          dateDeb,dateFin:TDateTime):integer;
function  DonneMailleSuivantDelai(const maille:char;
          dateDep:TDateTime;
          i:integer):hString;
function  ControleDateIntervalle(DateDeb,DateFin:TDateTime):boolean;     
function  DateAppartIntervalle(DateD,DateDeb,DateFin:TDateTime):boolean;
procedure DonneDateDebEtFinSaison(const Saison:hString;
          var dateDeb,dateFin:TDateTime);
procedure ChercheDateDDateFPeriode(const session:hString;
          var DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime);


function NbSemaineIntervalle(DateDeb,DateFin:TDateTime):integer; 
function NbMoisIntervalle(DateDeb,DateFin:TDateTime):integer;
function NbQuinzaineIntervalle(DateDeb,DateFin:TDateTime):integer;

procedure GrilleTailleNbTabTaille(const GrilleTaille:hString;
          var nb:integer;
          var TabTaille:array of hString);


procedure DatesDebutFinSaisonCmd(const SaiCmd:hString;var DateMin,DateMax:TDateTime);

function QuestBPInit(const codeSession:hString):boolean;
function QuestBPCalend(const codeSession:hString):boolean;

function QuestionBP():boolean;
 
function SessionBPInitialise(const session:hString):hString;

function NbTailleDansGrilleTaille(GrilleTaille:hString;
         var TabTai:array of hString):integer;

function DonneLibelleCodeAxe(const codestruct,num:hString):hString;
function DonneLibelleAxe(codeaxe:hString):hString;
function DonneLibelleCAxe(const codestruct,num,CodeAxe:hString):hString;
function ControleAxe(const codeAxe,structure,msg:hString):boolean;
function ControleValeurAxe(const codeAxe,valeurAxe,msg:hString):boolean;

procedure DonneNomTableChampsPrValeurAxe(const CodeAxe: hString;
          var NomTable,ChpCode,ChpLib,Chp2,Val2:hString);

function ChercheMaxCompteurPivot:integer;    

function ObjPrevConsultant():boolean;

function DonneCodeTabletteAxeStructure(const structure:hString):hString;
function ChercheTabletteValeurAxe(const codeAxe:hString):hString;

function  OkSessionLoiDefinie(const CodeSession:hString):boolean;

procedure ValRestrictionValaxe1234(valRestriction:hString;
          var valaxe1,valaxe2,valaxe3,valaxe4:hString);
function Valaxe1234ValRestriction(valaxe1,valaxe2,valaxe3,valaxe4:hString):hString;

function DonneLibelleDelai(Delai:TDateTime;BPInitialise:hString):hString;

function ChercheNumOrdreSession(const codeSession:hString):hString;

procedure ChercheAxeStructure(const structure:hString;
          var TabCodeAxe:array of hString;
          var NivMax:integer);

function AnalyseJoinSQL(const Join : hString; var TS, TS2, TS3: TstringList): hString;
function AnalyseJoinSQLPaie(const Join,TablePrincipale : hString): hString;

function BPTablettePlus(codeStruct: hString): hString;
function BPLibelleMaille(BPinitialise:hString):hString;

function AxeDeviseOK(const session: hString): boolean;
function SessionMinLoi(const session:hString):hString;
function CodeMarcheBiblioAxe : string;
function CodeMarcheCommun : string;
function VerifDateCube(Date : hString) : hString;

implementation

uses {$IFDEF PAIEGRH}EntPaie,{$ENDIF PAIEGRH}
     HEnt1,HMsgBox,CstCommun,UCtx,UCtxTrait,UUtil,BPUtil,DateUtils,
     UtilPgi;    // GC_20080331_GM_FQ_GC15917

function GestionBPTaille:boolean;
begin
 result:= DonneParamB(pb_GestTaille);
end;

function GestionBPMagasin:boolean;
begin
 result:= (DonneParamS(ps_GestMag) = '1');
end;

function GestionBPColoris:boolean;
begin
 result:= DonneParamB(pb_GestColoris);
end;

function GestionBPFS:boolean;
begin
 result := (DonneParamS(ps_GestFs) = '1') or (DonneParamS(ps_GestFs) = '0');
end;

//retourne vrai si Orli
function BPOkOrli:boolean;
begin
 {$IFNDEF BPORLI}
 result:=false;
 {$ELSE}
 result:=true;
 {$ENDIF}
end;

//****************************************************************************

function RemplaceCotteEspacePourSql(const S:hString):hString;
begin
 result:=trim(StringReplace(S,'"',' ',[rfReplaceAll]));
end;

//****************************************************************************

//procedure qui donne en résultat le code champ
//et le code left join pour la requete
//correspondant au code axe donné
procedure DonneCodeChpJoinPGI(const codeAxe,codeAs:hString;
          var chp,join:hString);
var Q:TQuery;
    joinX:hString;
begin
 chp:='';
 join:='';

 Q:=MOPenSql('SELECT QBX_NOMCHPREQUETE,QBX_LEFTJOINR FROM QBPBIBLIOAXE '+
             'WHERE QBX_CODEAXE="'+codeAxe+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
             'BPBasic (DonneCodeChpJoinPGI).',true);
 if not Q.eof
  then
   begin
    joinX:=Q.fields[1].asString;
    if joinX=''
     then chp:='[PREFIXE]'+Q.fields[0].asString
     else chp:=codeAs+'.[PREFIXE]'+Q.fields[0].asString;
    join:=StringReplace(joinX,'&',codeAs,[rfReplaceAll])
   end;
 ferme(Q);
end;

procedure ChercheCodeSqlChpAxeLibPGI(nivMax:integer;InitTree:boolean;
          TabCodeAxe:array of hString;
          var chpSql,joinSql:hString);
var j,nivTree:integer;
    table,chpcode,nomchp2,valchp2,codewhere,chpLib,nomRequete:hString;
    Q:TQuery;
begin
 chpSql:='';
 joinSql:='';
 nivTree := 0;
 if InitTree then
 begin
   nivTree := nivMax;
   nivMax := 1;
 end;

 for j:=1 to nivMax do
  begin
    if TabCodeAxe[j] = '' then
      continue;
   //code sql pour avoir le libelle
   Q:=MOpenSql('SELECT QBX_NOMTABLEAXE,QBX_CHPCODE,QBX_CHPLIB,'+
               'QBX_NOMCHP2,QBX_VALCHP2,QBX_NOMCHPREQUETE,QBX_LEFTJOINR '+
               'FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+TabCodeAxe[j]+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
               'BPBasic (ChercheCodeSqlChpAxeLib).',true);
   if not Q.eof
    then
     begin
      table:=Q.fields[0].asString;
      chpCode:=Q.fields[1].asString;
      chpLib:=Q.fields[2].asString;
      nomchp2:=Q.fields[3].asString;
      valchp2:=Q.fields[4].asString;
      if Q.fields[6].asString = '' then nomRequete:='[PREFIXE]'+Q.fields[5].asString
      else
      begin
        if InitTree then nomRequete:='Z'+IntToStr(nivTree)+'.[PREFIXE]'+Q.fields[5].asString
        else nomRequete:='Z'+IntToStr(j)+'.[PREFIXE]'+Q.fields[5].asString
      end;
      
      if chpSql<>'' then chpSql:=chpSql+' ,A'+IntToStr(j)+'.'+chpLib
      else chpSql:=' A'+IntToStr(j)+'.'+chpLib;

      { EVI / Correction : remise à '' du codewhere }
      if (nomchp2<>'') and (valchp2<>'') then codewhere:=' AND A'+IntToStr(j)+'.'+nomchp2+'="'+valchp2+'" '
      else codewhere:='';

      joinSql:=joinSql+' left join '+table+' A'+IntToStr(j)+
                    ' on A'+IntToStr(j)+'.'+chpcode+'='+nomRequete+' '+codewhere;
     end;
    ferme(Q);
  end;
 if chpsql<>''
  then chpsql:=','+chpsql;
end;

//procedure qui donne en résultat le code champ
//et le code left join pour la requete
//correspondant au code axe donné
procedure DonneCodeChpJoinPGIAs(const codeAxe:hString;
          var chp,join:hString);
var Q:TQuery;
begin
 chp:='';
 join:='';
 Q:=MOPenSql('SELECT QBX_NOMCHPREQUETE,QBX_LEFTJOINRAS FROM QBPBIBLIOAXE '+
             'WHERE QBX_CODEAXE="'+codeAxe+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
             'BPBasic (DonneCodeChpJoinPGIAs).',true);
 if not Q.eof
  then
   begin
    chp:='L1.'+Q.fields[0].asString;
    join:=Q.fields[1].asString;
   end;
 ferme(Q);
end;

procedure DonneCodeChpJoinPGINbNivAs(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,joins:hString);
var chp,join:hString;
    i:integer;
begin
 champs:='';
 joins:='';
 for i:=1 to niveauI do
  begin
   DonneCodeChpJoinPGIAs(TabCodeAxe[i],chp,join);
   if champs=''
    then champs:=chp
    else champs:=champs+','+chp+' ';

   if pos(joins,join)>0
    then joins:=join
    else joins:=joins+' '+join;
  end;
end;


procedure DonneCodeChpJoinPGINbNivLib(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,champsGB,joins:hString; const CodeSession: hString);
var chp,join,chpSS,joinSS:hString;
    i:integer;
    function SessionQte(): Boolean;
    var
      Q : TQuery;
    begin
      Result := False;
      Q := MOpenSql('SELECT QBS_VUEARBRE '+
                    'FROM QBPSESSIONBP WHERE QBS_CODESESSION="' + CodeSession + '"',
                    'BPBasic (DonneCodeChpJoinPGINbNivLib).', True);
      try
        if not Q.eof then
          Result := (Q.FindField('QBS_VUEARBRE').AsString = 'QTE');
      finally
        Ferme(Q);
      end;
    end;
begin
 champs:='';
 for i:=1 to niveauI do
  begin
   DonneCodeChpJoinPGI(TabCodeAxe[i],'Z'+IntToStr(i),chp,join);
   if champs=''
    then champs:=chp
    else champs:=champs+','+chp+' ';
   joinSS:='';
   chpSS:='';
   if i=niveauI then
     ChercheCodeSqlChpAxeLibPGI(niveauI,true,['',TabCodeAxe[i]],chpSS,joinSS);
   champs:=champs+RemplaceCotteEspacePourSql(chpSS);
   if pos(join,joins)>0
    then joins:=join+' '+joinSS
    else joins:=joins+' '+join+' '+joinSS;
  end;

  { EVI / Axe Devise
  if ctxCompta in V_PGI.PGIContexte then champs:=champs+',Y_DEVISE'
  else champs:=champs+',GL_DEVISE'; }
  champsGB:=champs;
  champs:=champs+',"' + V_PGI.DevisePivot + '"';
end;


//retourne le libellé d'une valeur axe
//suivant code axe et valeur axe donnés
function DonneLibelleValeurAxe(const CodeAxe,ValeurAxe: hString): hString;
var Q,Q2:TQuery;
    //mode,
    table,chpcode,chplib,nomchp2,valchp2,codewhere:hString;
begin
  result:='';
  table:='';
  if codeAxe='DELAI' then exit;
  if not BPOkOrli then
  begin
     Q:=MOpenSql('SELECT QBX_NOMTABLEAXE,QBX_CHPCODE,QBX_CHPLIB,'+
                'QBX_NOMCHP2,QBX_VALCHP2 '+
                'FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+codeaxe+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
                'BPBasic (DonneLibelleValeurAxe).',true);
     if not Q.eof
      then
       begin
        table:=Q.fields[0].asString;
        chpcode:=Q.fields[1].asString;
        chplib:=Q.fields[2].asString;
        nomchp2:=Q.fields[3].asString;
        valchp2:=Q.fields[4].asString;
        codewhere:='';
       end;
     ferme(Q);
     if table <>''
      then
       begin
        if (nomchp2<>'') and (valchp2<>'')
         then codewhere:=' AND '+nomchp2+'="'+valchp2+'" ';
        Q2:=MOPenSql('SELECT '+chplib+' FROM '+table+' WHERE '+
                     chpcode+'="'+ValeurAxe+'" '+codewhere,
                     'BPBasic (DonneLibelleValeurAxe).',true);
        if not Q2.eof
         then result:=Q2.fields[0].asString;
        ferme(Q2);
       end;
    end
   else
  //-----------------> ORLI
    begin
     Q:=MOPenSql('SELECT QBA_LIBVALEURAXE FROM QBPAXE WHERE '+
                 'QBA_CODEAXE="'+CodeAxe+'" AND QBA_VALEURAXE="'+ValeurAxe+'"',
                 'BPBasic (DonneLibelleValeurAxe).',true);
     if not Q.eof
      then result:=Q.fields[0].asString;
     ferme(Q);
    end;
  //ORLI <-----------------
end;

//recherche du codeaxe suivant le profil et le niveau
function DonneCodeAxeNiv(const session,Niv: hString): hString;
var Q: TQuery;
begin
  Result:='';
  Q:= openSql('SELECT QBS_CODEAXES'+Niv+' FROM QBPSESSIONBP WHERE '+
              'QBS_CODESESSION="'+session+'"',true);
  if not(Q.eof)
   then Result:=Q.fields[0].asString;
  ferme(Q);
end;

procedure ChercheCodeSqlChpAxeLib(nivMax:integer;
          TabCodeAxe:array of hString;
          var chpSql,joinSql:hString);
var j:integer;
    table,chpcode,nomchp2,valchp2,codewhere,chpLib:hString;
    Q:TQuery;
begin
 chpSql:='';
 joinSql:='';
// GC_20080117_DM_GC15713 DEBUT
//for j:=1 to nivMax-1 do
for j:=1 to nivMax do
// GC_20080117_DM_GC15713 FIN
  begin
    if TabCodeAxe[j] = '' then
      continue;
   //code sql pour avoir le libelle
   Q:=MOpenSql('SELECT QBX_NOMTABLEAXE,QBX_CHPCODE,QBX_CHPLIB,'+
               'QBX_NOMCHP2,QBX_VALCHP2 '+
               'FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+TabCodeAxe[j]+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
               'BPBasic (ChercheCodeSqlChpAxeLib).',true);
   if not Q.eof
    then
     begin
      table:=Q.fields[0].asString;
      chpCode:=Q.fields[1].asString;
      chpLib:=Q.fields[2].asString;
      nomchp2:=Q.fields[3].asString;
      valchp2:=Q.fields[4].asString;
      if table=''
       then
        begin
         if chpSql<>''
          then chpSql:=chpSql+',A'+IntToStr(j)+'.QBA_LIBVALEURAXE '
          else chpSql:=' A'+IntToStr(j)+'.QBA_LIBVALEURAXE ';
         if j=nivMax
          then joinSql:=joinSql+' left join QBPAXE A'+IntToStr(j)+
               ' on A'+IntToStr(j)+'.QBA_CODEAXE="'+TabCodeAxe[j]+
               '" AND A'+IntToStr(j)+'.QBA_VALEURAXE=QBR_VALEURAXE '
          else joinSql:=joinSql+' left join QBPAXE A'+IntToStr(j)+
               ' on A'+IntToStr(j)+'.QBA_CODEAXE="'+TabCodeAxe[j]+
               '" AND A'+IntToStr(j)+'.QBA_VALEURAXE=QBR_VALAXENIV'+IntToStr(j);
        end
       else
        begin
         if chpSql<>''
          then chpSql:=chpSql+' ,A'+IntToStr(j)+'.'+chpLib
          else chpSql:=' A'+IntToStr(j)+'.'+chpLib;
         if (nomchp2<>'') and (valchp2<>'')
          then codewhere:=' AND A'+IntToStr(j)+'.'+nomchp2+'="'+valchp2+'" '
          else codewhere:='' ;
         if j=nivMax
          then joinSql:=joinSql+' left join '+table+' A'+IntToStr(j)+
                  ' on A'+IntToStr(j)+'.'+chpcode+'=QBR_VALEURAXE '+codewhere
          else joinSql:=joinSql+' left join '+table+' A'+IntToStr(j)+
                  ' on A'+IntToStr(j)+'.'+chpcode+'=QBR_VALAXENIV'+IntToStr(j)
                  +codewhere;
        end;
     end;
    ferme(Q);
  end;
 if chpsql<>''
  then chpsql:=','+chpsql;
end;

procedure ChercheCodeSqlChpAxeLibAvecPivot(nivMax:integer;const codeAs:hString;
          TabCodeAxe,TabNumValAxe:array of hString;
          var chpSql,joinSql:hString);
var j:integer;
    table,chpcode,nomchp2,valchp2,codewhere,chpLib:hString;
    Q:TQuery;
begin
 chpSql:='';
 joinSql:='';
 for j:=1 to nivMax do
  begin
    if TabCodeAxe[j] = '' then
      continue;
   //code sql pour avoir le libelle
   Q:=MOpenSql('SELECT QBX_NOMTABLEAXE,QBX_CHPCODE,QBX_CHPLIB,'+
               'QBX_NOMCHP2,QBX_VALCHP2 '+
               'FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+TabCodeAxe[j]+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
               'BPBasic (ChercheCodeSqlChpAxeLibAvecPivot).',true);
   if not Q.eof
    then
     begin
      table:=Q.fields[0].asString;
      chpCode:=Q.fields[1].asString;
      chpLib:=Q.fields[2].asString;
      nomchp2:=Q.fields[3].asString;
      valchp2:=Q.fields[4].asString;
      if table=''
       then
        begin
         if chpSql<>''
          then chpSql:=chpSql+',A'+IntToStr(j)+'.QBA_LIBVALEURAXE '
          else chpSql:=' A'+IntToStr(j)+'.QBA_LIBVALEURAXE ';
         joinSql:=joinSql+' left join QBPAXE A'+IntToStr(j)+
               ' on A'+IntToStr(j)+'.QBA_CODEAXE="'+TabCodeAxe[j]+
               '" AND A'+IntToStr(j)+'.QBA_VALEURAXE='+codeAs+'QBV_VALAXEP'+TabNumValAxe[j];
        end
       else
        begin
         if chpSql<>''
          then chpSql:=chpSql+' ,A'+IntToStr(j)+'.'+chpLib
          else chpSql:=' A'+IntToStr(j)+'.'+chpLib;
         if (nomchp2<>'') and (valchp2<>'')
          then codewhere:=' AND A'+IntToStr(j)+'.'+nomchp2+'="'+valchp2+'" ';
         joinSql:=joinSql+' left join '+table+' A'+IntToStr(j)+
                  ' on A'+IntToStr(j)+'.'+chpcode+'='+codeAs+'QBV_VALAXEP'+TabNumValAxe[j]
                  +codewhere;
        end;
     end;
    ferme(Q);
  end;
 if chpsql<>''
  then chpsql:=','+chpsql;
end;


procedure DonneCodeChpJoinPGINbNiv(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs,joins:hString);
var chp,join:hString;
    i:integer;
begin
{ EVI / TRIM pour supprimer 3 espaces = (vide) = combo <<aucun>> }
 champs:='';
 joins:='';
 for i:=1 to niveauI do
  begin
   DonneCodeChpJoinPGI(TabCodeAxe[i],'Z'+IntToStr(i),chp,join);
   if champs=''
    then champs:='TRIM('+chp+')'
    else champs:=champs+',TRIM('+chp+') ';

   if pos(join,joins)>0
    then joins:=join
    else joins:=joins+' '+join;
  end;
end;

procedure DonneCodeChpUserRestrict(TabCodeAxe:array of hString;
          UserRestrict:hString; niveauI:integer;
          var UserSQLRestrict:hString);
var chp,join,values:hString;
    i:integer;
begin
 UserSQLRestrict:='';

 for i:=1 to niveauI do
  begin
   values := ReadTokenPipe(UserRestrict,'@@');
   DonneCodeChpJoinPGI(TabCodeAxe[i],'Z'+IntToStr(i),chp,join);
   if ((values <> '("")') AND (values <> '')) then UserSQLRestrict:=UserSQLRestrict+' AND '+chp+' IN '+values
  end;
end;


procedure DonneCodeChpJoinPGINbNivP(TabCodeAxe:array of hString;
          niveauI:integer;
          var champs:hString);
var chp,join:hString;
    i:integer;
begin
 champs:='';
 for i:=1 to niveauI do
  begin
   DonneCodeChpJoinPGI(TabCodeAxe[i],'Z'+IntToStr(i),chp,join);
   if champs=''
    then champs:=chp
    else champs:=champs+'||'+chp+' ';
  end;
end;

function CalculCoeffQte:boolean;
begin
 result:=(DonneParamS(ps_BPCoeffPerCAQte))='2';
end;

procedure MiseAJourTabletteValeurAAfficher;
begin
 if BPOkOrli
  then
  //-----------------> ORLI
   begin
    ExecuteSql('UPDATE COMMUN SET CO_ABREGE="X" WHERE CO_TYPE="QBV" AND '+
               '(CO_CODE="CCB" OR CO_CODE="CCN" OR CO_CODE="CEB" '+
               'OR CO_CODE="CQT" OR CO_CODE="RIE")');
    ExecuteSql('UPDATE COMMUN SET CO_ABREGE="-" WHERE CO_TYPE="MAR" AND '+
               '(CO_CODE="PAH" OR CO_CODE="PHT" OR CO_CODE="PTC" '+
               'OR CO_CODE="QTE" OR CO_CODE="UHT" OR CO_CODE="UTC")');
   end
  //ORLI <-----------------
  else
   begin
    ExecuteSql('UPDATE COMMUN SET CO_ABREGE="-" WHERE CO_TYPE="QBV" AND '+
               '(CO_CODE="CCB" OR CO_CODE="CCN" OR CO_CODE="CEB" '+
               'OR CO_CODE="CQT")');
    ExecuteSql('UPDATE COMMUN SET CO_ABREGE="" WHERE CO_TYPE="MAR" AND '+
               '(CO_CODE="PAH" OR CO_CODE="PHT" OR CO_CODE="PTC" '+
               'OR CO_CODE="QTE" OR CO_CODE="UHT" OR CO_CODE="UTC" OR CO_CODE="RIE")');
   end;
end;

procedure RechercheTabAff(var TabValAff,TabLibValAff:array of hString);
var val:hString;
    i:integer;
begin
  for i := 0 to 7 do { EVI / for i:=0 to 10 do --> ORLI }
  begin
   TabValAff[i]:='';
   TabLibValAff[i]:='';
  end;
 for i:=0 to 6 do
  begin
   Val:=DonneParamS(ps_BPValAff1+i);
   TabLibValAff[i]:=DonneParamS(ps_BPLibValAff1+i);
   if Val='MAR'
    then TabValAff[i]:='TAB7';
   if Val='PAH'
    then TabValAff[i]:='TAB6';
   if Val='UHT'
    then TabValAff[i]:='TAB5';
   if Val='UTC'
    then TabValAff[i]:='TAB4';
   if Val='PHT'
    then TabValAff[i]:='TAB3';
   if Val='QTE'
    then TabValAff[i]:='TAB2';
   if Val='PTC'
    then TabValAff[i]:='TAB1';
   if Val='CCB'
    then TabValAff[i]:='TAB1';
   if Val='CCN'
    then TabValAff[i]:='TAB3';
   if Val='CEB'
    then TabValAff[i]:='TAB4';
   if Val='CQT'
    then TabValAff[i]:='TAB2';         
  end;
end;


procedure ChangeControlChamp(const champ,codeControl:hString);
var i,j:integer;
begin
 i := PrefixeToNum(ExtractPrefixe(Champ));
 j := ChampToNum(Champ);
 if (i>-1) and (j>-1)
  then V_PGI.DEChamps[i, j].Control:=codeControl;
end;

procedure ChangeLibelleChamp(const champ,libelle:hString);
var i,j:integer;
begin
 i := PrefixeToNum(ExtractPrefixe(Champ));
 j := ChampToNum(Champ);
 if (i>-1) and (j>-1) then V_PGI.DEChamps[i, j].Libelle := TraduireMemoire(libelle);
end;

procedure ShowHideChamp(ValAff : integer);
var i,nbVal :integer;
TabValAff : array[1..7] of hstring;
begin
  for i := 1 to 7 do
  begin
    if ValAff = i then TabValAff[i]:='L' else TabValAff[i]:='';
  end;

  //champs à afficher suivant le paramètre liste des valeurs à afficher
  for i:=1 to 7 do
  begin
    nbVal := i;
    if i<>2 then
    begin
      if nbVal<>1 then nbVal:=nbVal-1;
      ChangeControlChamp('QBR_REF'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_REFPRCT'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_OP'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_OPPRCT'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_SAISI'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_EVOLVAL'+IntToStr(nbVal),TabValAff[i]);
      ChangeControlChamp('QBR_EVOLPRCT'+IntToStr(nbVal),TabValAff[i]);
    end
    else
    begin
      ChangeControlChamp('QBR_QTEREF',TabValAff[i]);
      ChangeControlChamp('QBR_QTEREFPRCT',TabValAff[i]);
      ChangeControlChamp('QBR_QTEC',TabValAff[i]);
      ChangeControlChamp('QBR_QTECPRCT',TabValAff[i]);
      ChangeControlChamp('QBR_SAISIQTE',TabValAff[i]);
      ChangeControlChamp('QBR_EVOLVALQTE',TabValAff[i]);
      ChangeControlChamp('QBR_EVOLPRCTQTE',TabValAff[i]);
    end;
  end
end;


procedure ChangeListeLibelleColonnes(OkObjectif:boolean;Session:string;TabValAff,TabLibValAff:array of hString);
var codeOP,nbVal:hString;
    i,ValeurAffiche : integer;
    TabValAffPaie: array [0..6] of hString;
begin
  codeOP:='Prévision';
  if OkObjectif then CodeOP:='Objectif';

  if not BPOkOrli then
  begin
    //Détermine les champs disponibles selon la valeur affichée
    ValeurAffiche := 0;
    if TabValAff[0]='PTC' then ValeurAffiche := 1;
    if TabValAff[0]='QTE' then ValeurAffiche := 2;
    if TabValAff[0]='PHT' then ValeurAffiche := 3;
    if TabValAff[0]='UTC' then ValeurAffiche := 4;
    if TabValAff[0]='UHT' then ValeurAffiche := 5;
    if TabValAff[0]='PAH' then ValeurAffiche := 6;
    if TabValAff[0]='MAR' then ValeurAffiche := 7;

    Case ContextBP of
      0,1 : begin     //GC-Mode
            //------------------------------------------------------------------------

            //change les libellés pour liste concernant fiche arbre
            //CA TTC net
            ChangeLibelleChamp('QBR_REF1','Valeurs réference (CA TTC net)');
            ChangeLibelleChamp('QBR_REFPRCT1','% référence (CA TTC net)');
            ChangeLibelleChamp('QBR_OP1',CodeOP+' (CA TTC net)');
            ChangeLibelleChamp('QBR_OPPRCT1','Nouveau % répartition (CA TTC net)');
            ChangeLibelleChamp('QBR_SAISI1',CodeOP+' saisi (CA TTC net)');
            ChangeLibelleChamp('QBR_EVOLVAL1','Montant évolution (CA TTC net)');
            ChangeLibelleChamp('QBR_EVOLPRCT1','% évolution (CA TTC net)');
            //CA HT net
            ChangeLibelleChamp('QBR_REF2','Valeur réference (CA HT net)');
            ChangeLibelleChamp('QBR_REFPRCT2','% référence (CA HT net)');
            ChangeLibelleChamp('QBR_OP2',CodeOP+' (CA HT net)');
            ChangeLibelleChamp('QBR_OPPRCT2','Nouveau % répartition (CA HT net)');
            ChangeLibelleChamp('QBR_SAISI2',CodeOP+' saisi (CA HT net)');
            ChangeLibelleChamp('QBR_EVOLVAL2','Montant évolution (CA HT net)');
            ChangeLibelleChamp('QBR_EVOLPRCT2','% évolution (CA HT net)');
            //CA HT avant remise
            ChangeLibelleChamp('QBR_REF3','Valeur réference (CA HT avant remise)');
            ChangeLibelleChamp('QBR_REFPRCT3','% référence (CA HT avant remise)');
            ChangeLibelleChamp('QBR_OP3',CodeOP+' (CA HT avant remise)');
            ChangeLibelleChamp('QBR_OPPRCT3','Nouveau % répartition (CA HT avant remise)');
            ChangeLibelleChamp('QBR_SAISI3',CodeOP+' saisi (CA HT avant remise)');
            ChangeLibelleChamp('QBR_EVOLVAL3','Montant évolution (CA HT avant remise)');
            ChangeLibelleChamp('QBR_EVOLPRCT3','% évolution (CA HT avant remise)');
            //CA TTC avant remise
            ChangeLibelleChamp('QBR_REF4','Valeur réference (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_REFPRCT4','% référence (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_OP4',CodeOP+' (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_OPPRCT4','Nouveau % répartition (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_SAISI4',CodeOP+' saisi (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_EVOLVAL4','Montant évolution (CA TTC avant remise)');
            ChangeLibelleChamp('QBR_EVOLPRCT4','% évolution (CA TTC avant remise)');
            //Cumul achat en PAMP
            ChangeLibelleChamp('QBR_REF5','Valeur réference (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_REFPRCT5','% référence (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_OP5',CodeOP+' (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_OPPRCT5','Nouveau % répartition (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_SAISI5',CodeOP+' saisi (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_EVOLVAL5','Montant évolution (Cumul achat en PAMP)');
            ChangeLibelleChamp('QBR_EVOLPRCT5','% évolution (Cumul achat en PAMP)');
            //Marge
            ChangeLibelleChamp('QBR_REF6','Valeur réference (Marge)');
            ChangeLibelleChamp('QBR_REFPRCT6','% référence (Marge)');
            ChangeLibelleChamp('QBR_OP6',CodeOP+' (Marge)');
            ChangeLibelleChamp('QBR_OPPRCT6','Nouveau % répartition (Marge)');
            ChangeLibelleChamp('QBR_SAISI6',CodeOP+' saisi (Marge)');
            ChangeLibelleChamp('QBR_EVOLVAL6','Montant évolution (Marge)');
            ChangeLibelleChamp('QBR_EVOLPRCT6','% évolution (Marge)');
            //Quantité
            ChangeLibelleChamp('QBR_QTEREF','Valeur réference (Quantité)');
            ChangeLibelleChamp('QBR_QTEREFPRCT','% référence (Quantité)');
            ChangeLibelleChamp('QBR_QTEC',CodeOP+' (Quantité)');
            ChangeLibelleChamp('QBR_QTECPRCT','Nouveau % répartition (Quantité)');
            ChangeLibelleChamp('QBR_SAISIQTE',CodeOP+' saisi (Quantité)');
            ChangeLibelleChamp('QBR_EVOLVALQTE','Montant évolution (Quantité)');
            ChangeLibelleChamp('QBR_EVOLPRCTQTE','% évolution (Quantité)');

            ChangeLibelleChamp('QBS_NATURECMD','Nature de documents');
            ShowHideChamp(ValeurAffiche);

            //------------------------------------------------------------------------
          end;
      2 : begin  //Compta
            //------------------------------------------------------------------------

            //change les libellés pour liste concernant fiche arbre
            //Débit - Crédit
            ChangeLibelleChamp('QBR_REF1','Valeurs réference (Débit - Crédit)');
            ChangeLibelleChamp('QBR_REFPRCT1','% référence (Débit - Crédit)');
            ChangeLibelleChamp('QBR_OP1',CodeOP+' (Débit - Crédit)');
            ChangeLibelleChamp('QBR_OPPRCT1','Nouveau % répartition (Débit - Crédit)');
            ChangeLibelleChamp('QBR_SAISI1',CodeOP+' saisi (Débit - Crédit)');
            ChangeLibelleChamp('QBR_EVOLVAL1','Montant évolution (Débit - Crédit)');
            ChangeLibelleChamp('QBR_EVOLPRCT1','% évolution (Débit - Crédit)');
            //Crédit - Débit
            ChangeLibelleChamp('QBR_REF2','Valeur réference (Crédit - Débit)');
            ChangeLibelleChamp('QBR_REFPRCT2','% référence (Crédit - Débit)');
            ChangeLibelleChamp('QBR_OP2',CodeOP+' (Crédit - Débit)');
            ChangeLibelleChamp('QBR_OPPRCT2','Nouveau % répartition (Crédit - Débit)');
            ChangeLibelleChamp('QBR_SAISI2',CodeOP+' saisi (Crédit - Débit)');
            ChangeLibelleChamp('QBR_EVOLVAL2','Montant évolution (Crédit - Débit)');
            ChangeLibelleChamp('QBR_EVOLPRCT2','% évolution (Crédit - Débit)');

            ChangeLibelleChamp('QBS_NATURECMD','Nature de documents');
            ShowHideChamp(ValeurAffiche);

            //------------------------------------------------------------------------
          end;
      3 : begin //Paie
            //------------------------------------------------------------------------

            //change les libellés pour liste concernant fiche arbre
            LibValAff(Session,TabValAffPaie);

            //CA TTC net
            ChangeLibelleChamp('QBR_REF1','Valeurs réference ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_REFPRCT1','% référence ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_OP1',CodeOP+' ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_OPPRCT1','Nouveau % répartition ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_SAISI1',CodeOP+' saisi ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_EVOLVAL1','Montant évolution ('+TabValAffPaie[1]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT1','% évolution ('+TabValAffPaie[1]+')');
            //CA HT net
            ChangeLibelleChamp('QBR_REF2','Valeur réference ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_REFPRCT2','% référence ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_OP2',CodeOP+' ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_OPPRCT2','Nouveau % répartition ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_SAISI2',CodeOP+' saisi ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_EVOLVAL2','Montant évolution ('+TabValAffPaie[2]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT2','% évolution ('+TabValAffPaie[2]+')');
            //CA HT avant remise
            ChangeLibelleChamp('QBR_REF3','Valeur réference ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_REFPRCT3','% référence ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_OP3',CodeOP+' ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_OPPRCT3','Nouveau % répartition ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_SAISI3',CodeOP+' saisi ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_EVOLVAL3','Montant évolution ('+TabValAffPaie[3]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT3','% évolution ('+TabValAffPaie[3]+')');
            //CA TTC avant remise
            ChangeLibelleChamp('QBR_REF4','Valeur réference ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_REFPRCT4','% référence ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_OP4',CodeOP+' ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_OPPRCT4','Nouveau % répartition ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_SAISI4',CodeOP+' saisi ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_EVOLVAL4','Montant évolution ('+TabValAffPaie[4]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT4','% évolution ('+TabValAffPaie[4]+')');
            //Cumul achat en PAMP
            ChangeLibelleChamp('QBR_REF5','Valeur réference ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_REFPRCT5','% référence ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_OP5',CodeOP+' ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_OPPRCT5','Nouveau % répartition ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_SAISI5',CodeOP+' saisi ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_EVOLVAL5','Montant évolution ('+TabValAffPaie[5]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT5','% évolution ('+TabValAffPaie[5]+')');
            //Marge
            ChangeLibelleChamp('QBR_REF6','Valeur réference ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_REFPRCT6','% référence ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_OP6',CodeOP+' ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_OPPRCT6','Nouveau % répartition ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_SAISI6',CodeOP+' saisi ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_EVOLVAL6','Montant évolution ('+TabValAffPaie[6]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT6','% évolution ('+TabValAffPaie[6]+')');
            //Quantité
            ChangeLibelleChamp('QBR_QTEREF','Valeur réference ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_QTEREFPRCT','% référence ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_QTEC',CodeOP+' ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_QTECPRCT','Nouveau % répartition ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_SAISIQTE',CodeOP+' saisi ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_EVOLVALQTE','Montant évolution ('+TabValAffPaie[0]+')');
            ChangeLibelleChamp('QBR_EVOLPRCTQTE','% évolution ('+TabValAffPaie[0]+')');

            //ChangeLibelleChamp('QBS_NATURECMD','Nature de documents');
            //ShowHideChamp(ValeurAffiche);

            //------------------------------------------------------------------------
          end

    end;

    //change libellé et colonnes affichées pour fiche session
    ChangeLibelleChamp('QBS_DATEDEBC','Date début (période objectif)');
    ChangeLibelleChamp('QBS_DATEFINC','Date fin (période objectif)');

    { EVI / Axe Devise }
    ChangeControlChamp('QBR_DEVISE','');
    { EVI / Suppression des axes propres à ORLI }
    ChangeControlChamp('QBS_CODESTRUCT','');
    ChangeControlChamp('QBS_CODEAXES6','');
    ChangeControlChamp('QBS_CODEAXES7','');
    ChangeControlChamp('QBS_CODEAXES8','');
    ChangeControlChamp('QBS_CODEAXES9','');
    ChangeControlChamp('QBS_CODEAXES10','');
    ChangeControlChamp('QBS_EXTRAPOLABLE','');
    ChangeControlChamp('QBS_INITCOEFF','');
    ChangeControlChamp('QBS_INITPREVISION','');
    ChangeControlChamp('QBS_MODEGESTPREV','');
    ChangeControlChamp('QBS_SAISONCMDC','');
    ChangeControlChamp('QBS_SAISONCMDREF','');
    ChangeControlChamp('QBS_SESIONOBJECTIF','');
    ChangeControlChamp('QBS_SESSIONFERMEE','');
    ChangeControlChamp('QBS_SESSIONNETTE','');
    ChangeControlChamp('QBS_SESSIONBRUTE','');
    ChangeControlChamp('QBS_TYPENATURE','');
    ChangeControlChamp('QBS_SESSIONECLATAI','');
    ChangeControlChamp('QBS_BPINITIALISE','');
    ChangeControlChamp('QBS_CODEAXE1','');
    ChangeControlChamp('QBS_CODEAXE2','');
    ChangeControlChamp('QBS_CODEAXE3','');
    ChangeControlChamp('QBS_CODEAXE4','');
    ChangeControlChamp('QBS_VALEURAXE1','');
    ChangeControlChamp('QBS_VALEURAXE2','');
    ChangeControlChamp('QBS_VALEURAXE3','');
    ChangeControlChamp('QBS_VALEURAXE4','');
    ChangeControlChamp('QBR_CARETENU','');
    ChangeControlChamp('QBR_COEFFCALCUL','');
    ChangeControlChamp('QBR_COEFFGENE','');
    ChangeControlChamp('QBR_COEFFRETENU','');
    ChangeControlChamp('QBR_EVOLPREVCA','');
    ChangeControlChamp('QBR_EVOLPREVCAPRCT','');
    ChangeControlChamp('QBR_EVOLPREVQTE','');
    ChangeControlChamp('QBR_EVOLPREVQTEPRC','');
    ChangeControlChamp('QBR_EXTRAPOLABLE','');
    ChangeControlChamp('QBR_HISTO','');
    ChangeControlChamp('QBR_HISTOCA','');
    ChangeControlChamp('QBR_NBCLTC','');
    ChangeControlChamp('QBR_NBCLTNOUVEAU','');
    ChangeControlChamp('QBR_NBCLTPERDU','');
    ChangeControlChamp('QBR_NBCLTPROSPECT','');
    ChangeControlChamp('QBR_NBCLTREF','');
    ChangeControlChamp('QBR_NBCLTRESTAVOIR','');
    ChangeControlChamp('QBR_NBCLTVU','');
    ChangeControlChamp('QBR_NEWPROSPECT','');
    ChangeControlChamp('QBR_NOUVEAU','');
    ChangeControlChamp('QBR_NOUVEAUQTE','');
    ChangeControlChamp('QBR_PERDU','');
    ChangeControlChamp('QBR_PERDUQTE','');
    ChangeControlChamp('QBR_PREVU','');
    ChangeControlChamp('QBR_PREVUCA','');
    ChangeControlChamp('QBR_PROSPECT','');
    ChangeControlChamp('QBR_PROSPECTQTE','');
    ChangeControlChamp('QBR_QTERETENUE','');
    ChangeControlChamp('QBR_QTET1','');
    ChangeControlChamp('QBR_QTET2','');
    ChangeControlChamp('QBR_QTET3','');
    ChangeControlChamp('QBR_QTET4','');
    ChangeControlChamp('QBR_QTET5','');
    ChangeControlChamp('QBR_QTET6','');
    ChangeControlChamp('QBR_QTET7','');
    ChangeControlChamp('QBR_QTET8','');
    ChangeControlChamp('QBR_QTET9','');
    ChangeControlChamp('QBR_QTET10','');
    ChangeControlChamp('QBR_QTET11','');
    ChangeControlChamp('QBR_QTET12','');
    ChangeControlChamp('QBR_QTET13','');
    ChangeControlChamp('QBR_QTET14','');
    ChangeControlChamp('QBR_QTET15','');
    ChangeControlChamp('QBR_QTET16','');
    ChangeControlChamp('QBR_QTET17','');
    ChangeControlChamp('QBR_QTET18','');
    ChangeControlChamp('QBR_QTET19','');
    ChangeControlChamp('QBR_QTET20','');
    ChangeControlChamp('QBR_REALISE','');
    ChangeControlChamp('QBR_REALISECA','');
    ChangeControlChamp('QBR_RESTEAVISTER','');
    ChangeControlChamp('QBR_SAISIPREVCA','');
    ChangeControlChamp('QBR_SAISIPREVQTE','');
    ChangeControlChamp('QBR_TOTALCOURANT','');
    ChangeControlChamp('QBR_TOTALREF','');
    ChangeControlChamp('QBR_VALMODIF','');
    ChangeControlChamp('QBR_VUCOURANT','');
    ChangeControlChamp('QBR_VUCOURANTQTE','');
    ChangeControlChamp('QBR_VUREF','');
    ChangeControlChamp('QBR_VUREFQTE','');
    ChangeControlChamp('QBR_XVISIT','');

  end;

 if BPOkOrli
  then
  //-----------------> ORLI
   begin
    //------------------------------------------------------------------------

    //change libellé et colonnes affichées pour fiche session
    ChangeLibelleChamp('QBS_DATEDEBC','Date début (période objectif)');
    ChangeLibelleChamp('QBS_DATEFINC','Date fin (période objectif)');   
    ChangeLibelleChamp('QBS_NATURECMD','Nature de documents');

    ChangeControlChamp('QBS_CODESTRUCT','');
    ChangeControlChamp('QBS_CODEAXES6','');
    ChangeControlChamp('QBS_CODEAXES7','');
    ChangeControlChamp('QBS_CODEAXES8','');
    ChangeControlChamp('QBS_CODEAXES9','');
    ChangeControlChamp('QBS_CODEAXES10','');
    ChangeControlChamp('QBS_EXTRAPOLABLE','');
    ChangeControlChamp('QBS_INITCOEFF','');
    ChangeControlChamp('QBS_INITPREVISION','');
    ChangeControlChamp('QBS_MODEGESTPREV','');
    ChangeControlChamp('QBS_SAISONCMDC','');
    ChangeControlChamp('QBS_SAISONCMDREF','');
    ChangeControlChamp('QBS_SESIONOBJECTIF','');
    ChangeControlChamp('QBS_SESSIONFERMEE','');
    ChangeControlChamp('QBS_SESSIONNETTE','');
    ChangeControlChamp('QBS_SESSIONBRUTE','');
    ChangeControlChamp('QBS_TYPENATURE','');
    ChangeControlChamp('QBS_SESSIONECLATAI','');
    ChangeControlChamp('QBS_BPINITIALISE','');
    ChangeControlChamp('QBS_CODEAXE1','');   
    ChangeControlChamp('QBS_CODEAXE2','');
    ChangeControlChamp('QBS_CODEAXE3','');
    ChangeControlChamp('QBS_CODEAXE4','');
    ChangeControlChamp('QBS_VALEURAXE1','');
    ChangeControlChamp('QBS_VALEURAXE2','');
    ChangeControlChamp('QBS_VALEURAXE3','');
    ChangeControlChamp('QBS_VALEURAXE4','');

    //CA 1
    ChangeControlChamp('QBR_REF1','');
    ChangeControlChamp('QBR_REFPRCT1','');
    ChangeControlChamp('QBR_OP1','');
    ChangeControlChamp('QBR_OPPRCT1','');
    ChangeControlChamp('QBR_SAISI1','');
    ChangeControlChamp('QBR_EVOLVAL1','');
    ChangeControlChamp('QBR_EVOLPRCT1','');
    //CA 2
    ChangeControlChamp('QBR_REF2','');
    ChangeControlChamp('QBR_REFPRCT2','');
    ChangeControlChamp('QBR_OP2','');
    ChangeControlChamp('QBR_OPPRCT2','');
    ChangeControlChamp('QBR_SAISI2','');
    ChangeControlChamp('QBR_EVOLVAL2','');
    ChangeControlChamp('QBR_EVOLPRCT2','');
    //CA 3
    ChangeControlChamp('QBR_REF3','');
    ChangeControlChamp('QBR_REFPRCT3','');
    ChangeControlChamp('QBR_OP3','');
    ChangeControlChamp('QBR_OPPRCT3','');
    ChangeControlChamp('QBR_SAISI3','');
    ChangeControlChamp('QBR_EVOLVAL3','');
    ChangeControlChamp('QBR_EVOLPRCT3','');
    //CA 4
    ChangeControlChamp('QBR_REF4','');
    ChangeControlChamp('QBR_REFPRCT4','');
    ChangeControlChamp('QBR_OP4','');
    ChangeControlChamp('QBR_OPPRCT4','');
    ChangeControlChamp('QBR_SAISI4','');
    ChangeControlChamp('QBR_EVOLVAL4','');
    ChangeControlChamp('QBR_EVOLPRCT4','');
    //CA 5
    ChangeControlChamp('QBR_REF5','');
    ChangeControlChamp('QBR_REFPRCT5','');
    ChangeControlChamp('QBR_OP5','');
    ChangeControlChamp('QBR_OPPRCT5','');
    ChangeControlChamp('QBR_SAISI5','');
    ChangeControlChamp('QBR_EVOLVAL5','');
    ChangeControlChamp('QBR_EVOLPRCT5','');
    //CA 6
    ChangeControlChamp('QBR_REF6','');
    ChangeControlChamp('QBR_REFPRCT6','');
    ChangeControlChamp('QBR_OP6','');
    ChangeControlChamp('QBR_OPPRCT6','');
    ChangeControlChamp('QBR_SAISI6','');
    ChangeControlChamp('QBR_EVOLVAL6','');
    ChangeControlChamp('QBR_EVOLPRCT6','');    
    //Qte
    ChangeControlChamp('QBR_QTEREF','');
    ChangeControlChamp('QBR_QTEREFPRCT','');
    ChangeControlChamp('QBR_QTEC','');
    ChangeControlChamp('QBR_QTECPRCT','');
    ChangeControlChamp('QBR_SAISIQTE','');
    ChangeControlChamp('QBR_EVOLVALQTE','');
    ChangeControlChamp('QBR_EVOLPRCTQTE','');

    //champs à afficher suivant le paramètre liste des valeurs à afficher
    for i:=0 to 7 do
     begin
      if TabValAff[i]<>''
       then
        begin
         nbVal:=copy(TabValAff[i],4,1);
         if nbVal<>'2'
          then
           begin
            if nbVal<>'1'
             then nbVal:=IntToStr(VALEURI(nbVal)-1);

            ChangeLibelleChamp('QBR_REF'+nbVal,'Valeur réference ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_REFPRCT'+nbVal,'% référence ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_OP'+nbVal,CodeOP+' ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_OPPRCT'+nbVal,'Nouveau % répartition ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_SAISI'+nbVal,CodeOP+' saisi ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_EVOLVAL'+nbVal,'Montant évolution ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_EVOLPRCT'+nbVal,'% évolution ('+TabLibValAff[i]+')');

            ChangeControlChamp('QBR_REF'+nbVal,'L');
            ChangeControlChamp('QBR_REFPRCT'+nbVal,'L');
            ChangeControlChamp('QBR_OP'+nbVal,'L');
            ChangeControlChamp('QBR_OPPRCT'+nbVal,'L');
            ChangeControlChamp('QBR_SAISI'+nbVal,'L');
            ChangeControlChamp('QBR_EVOLVAL'+nbVal,'L');
            ChangeControlChamp('QBR_EVOLPRCT'+nbVal,'L');
           end
          else
           begin                                        
            ChangeLibelleChamp('QBR_QTEREF','Valeur réference ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_QTEREFPRCT','% référence ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_QTEC',CodeOP+' ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_QTECPRCT','Nouveau % répartition ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_SAISIQTE',CodeOP+' saisi ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_EVOLVALQTE','Montant évolution ('+TabLibValAff[i]+')');
            ChangeLibelleChamp('QBR_EVOLPRCTQTE','% évolution ('+TabLibValAff[i]+')');

            ChangeControlChamp('QBR_QTEREF','L');
            ChangeControlChamp('QBR_QTEREFPRCT','L');
            ChangeControlChamp('QBR_QTEC','L');
            ChangeControlChamp('QBR_QTECPRCT','L');
            ChangeControlChamp('QBR_SAISIQTE','L');
            ChangeControlChamp('QBR_EVOLVALQTE','L');
            ChangeControlChamp('QBR_EVOLPRCTQTE','L');
           end;
        end;
     end;
    //------------------------------------------------------------------------
   end;
  //ORLI <-----------------


end;



procedure DonneNumTrimestre(DateT:TDateTime;var NumTrimestre,Annee:integer);
var moisT,anT,JourT:word;
begin
 DecodeDate(dateT,AnT,moisT,JourT);
 case moisT of
  1,2,3:NumTrimestre:=1;
  4,5,6:NumTrimestre:=2;
  7,8,9:NumTrimestre:=3;
  10,11,12:NumTrimestre:=4;
 end;
 Annee:=AnT;
end;

procedure DonneNumQuadrimestre(DateT:TDateTime;var NumQuadrimestre,Annee:integer);
var moisT,anT,JourT:word;
begin  
 DecodeDate(dateT,AnT,moisT,JourT);
 case moisT of
  1,2,3,4:NumQuadrimestre:=1;
  5,6,7,8:NumQuadrimestre:=2;
  9,10,11,12:NumQuadrimestre:=3;
 end;
 Annee:=AnT;
end;


procedure DonneDateDebFinTrimestre(NumT,AnT:integer;var DateDeb,DateFin:TDateTime);
begin
 case numT of
  1 : begin
       DateDeb:=EncodeDate(AnT,1,1);
       DateFin:=EncodeDate(AnT,3,31);
      end;
  2 : begin
       DateDeb:=EncodeDate(AnT,4,1);
       DateFin:=EncodeDate(AnT,6,30);
      end;
  3 : begin
       DateDeb:=EncodeDate(AnT,7,1);
       DateFin:=EncodeDate(AnT,9,30);
      end;
  4 : begin
       DateDeb:=EncodeDate(AnT,10,1);
       DateFin:=EncodeDate(AnT,12,31);
      end;
 end;
end;

procedure DonneDateDebFinQuadrimestre(NumT,AnT:integer;var DateDeb,DateFin:TDateTime);
begin
 case numT of
  1 : begin
       DateDeb:=EncodeDate(AnT,1,1);
       DateFin:=EncodeDate(AnT,4,30);
      end;
  2 : begin
       DateDeb:=EncodeDate(AnT,5,1);
       DateFin:=EncodeDate(AnT,8,31);
      end;
  3 : begin
       DateDeb:=EncodeDate(AnT,9,1);
       DateFin:=EncodeDate(AnT,12,31);
      end;
 end;
end;


function DonneNbIntervalleMailleDateDebDateFin(const maille:hString;
         dateDeb,dateFin:TDateTime):integer;
var anDeb,moisdeb,JourDeb:word;
    anFin,moisFin,JourFin:word;
    diff,numsemF,numsemD,aDeb,aFin:integer;   
    perD,perF: TperiodeP;
    NumTD,AnTD,NumTF,AnTF:integer;
begin
 diff:=0;

 DecodeDate(dateDeb,anDeb,moisdeb,JourDeb);
 DecodeDate(dateFin,anFin,moisFin,JourFin);
 numsemD:=numSemaine(dateDeb,aDeb);
 numsemF:=numSemaine(dateFin,aFIn);

 if maille='2'
  then
   begin
    if aFin=aDeb
     then diff:=numsemF-numsemD+1
     else diff:=(52-numsemD)+numsemF;
    if diff>=53
     then diff:=52;
   end;
 if maille='3'
  then
   begin  
    diff:=NbQuinzaineIntervalle(DateDeb,DateFin);
   end;
 if maille='4'
  then
   begin
    if anFin=anDeb
     then diff:=moisFin-moisDeb+1
     else diff:=(12-moisDeb)+moisFin;
   end;
 if maille='5'
  then
   begin
    PerD:=FindPeriodeP(DateDeb);
    PerF:=FindPeriodeP(DateFin);
    diff:=PerF.periode-PerD.periode+1;
   end;
 if maille='6'
  then
   begin
    DonneNumTrimestre(DateDeb,NumTD,AnTD);
    DonneNumTrimestre(DateFin,NumTF,AnTF);
    if AnTD=AnTF
     then diff:=NumTF-NumTD+1
     else diff:=(4-NumTD)+NumTF;
   end;
 if maille='7'
  then
   begin
    DonneNumQuadrimestre(DateDeb,NumTD,AnTD);
    DonneNumQuadrimestre(DateFin,NumTF,AnTF);
    if AnTD=AnTF
     then diff:=NumTF-NumTD+1
     else diff:=(3-NumTD)+NumTF;
   end;
 result:=diff;
end;

function NbQuinzaineIntervalle(DateDeb,DateFin:TDateTime):integer;
var anDeb,moisDeb,anF,moisF:word;
    QuinzDeb,NumQuinzDeb,QuinzF,NumQuinzF:word;
begin
 result:=0;
 Date_AnNumQuinz(DateDeb,anDeb,moisDeb,QuinzDeb,NumQuinzDeb);
 Date_AnNumQuinz(DateFin,anF,moisF,QuinzF,NumQuinzF);
 if anDeb=anF
  then result:=NumQuinzF-NumQuinzDeb+1;
 if anF>anDeb
  then result:=(24-NumQuinzDeb)+NumQuinzF+1;
end;

procedure GrilleTailleNbTabTaille(const GrilleTaille:hString;
          var nb:integer;
          var TabTaille:array of hString);
var Q:TQuery;
    i:integer;
begin
 nb:=0;
 Q:=MOpenSql('SELECT QTA_TAILLE1,QTA_TAILLE2,QTA_TAILLE3,QTA_TAILLE4,QTA_TAILLE5,' +
             'QTA_TAILLE6,QTA_TAILLE7,QTA_TAILLE8,QTA_TAILLE9,QTA_TAILLE10,' +
             'QTA_TAILLE11,QTA_TAILLE12,QTA_TAILLE13,QTA_TAILLE14,QTA_TAILLE15,' +
             'QTA_TAILLE16,QTA_TAILLE17,QTA_TAILLE18,QTA_TAILLE19,QTA_TAILLE20 ' +
             ' FROM QTAILLE WHERE '+WhereCtx(['qtaille'])+
             ' AND QTA_CODETAILLE="'+GrilleTaille+'"',
             'GrilleTailleNbTabTaille (BPBasic).',true);
 if not Q.eof
  then
   begin
     for i:=0 to 19 do
      begin
       TabTaille[i]:=Q.fields[i].asString;
       if TabTaille[i]<>''
        then nb:=nb+1;
      end;
   end;
 ferme(Q);
end;


procedure DatesDebutFinSaisonCmd(const SaiCmd:hString;var DateMin,DateMax:TDateTime);
var Q:TQuery;
begin
 Q:=MOpenSql('SELECT MIN(QBV_DATEPIVOT),MAX(QBV_DATEPIVOT) FROM QBPPIVOT '+
             'WHERE QBV_VALAXEP8="'+SaiCmd+'"',
             'BPBasic (DatesDebutFinSaisonCmd).',true);
 if not Q.eof
  then
   begin
    DateMin:=Q.Fields[0].asDateTime;
    DateMax:=Q.Fields[1].asDateTime;
   end;
 ferme(Q);
end;

function DonneMailleSuivantDelai(const maille:char;
         dateDep:TDateTime;
         i:integer):hString;
var DelaiMaille:hString;
    anDep,moisdep,JourDep:word;
    code, AnSemaine:hString;
    PremJDep,DateI:TDateTime;
    per: TperiodeP;
    function TrouveAnnee(Const LaDate: tDateTime): word;
    var
      y, m, d: word;
    begin
      Decodedate(LaDate, y, m, d);
      Result := y;
    end;
begin
 DelaiMaille:=DonneParamS(ps_BPDelaiMaille);
 DecodeDate(dateDep,anDep,moisdep,JourDep);
 if not BPOkOrli
  then
   begin
    case maille of
     '2' : begin

            DateI := DateDep + (i*7);
            AnSemaine := DonneAnSemaine(DateI);
            PremJDep:=PremierJourSemaine(ValeurI(Copy(AnSemaine,4,2)), an2_an4(ValeurI(Copy(AnSemaine,1,2))));
            //resultat
            code:=DateTimeToStr(PremJDep);
           end;
     '3' : begin
            DateI:=PLUSMOIS(DateDep,i);
            code:=DateTimeToStr(DEBUTDEMOIS(DateI));
           end;
     '5' : begin
            DateI:=PLUSMOIS(DateDep,i);
            code:=DateTimeToStr(DEBUTDEMOIS(DateI)+15);
           end;
     '4' : begin
            //mois
            DateI:=PLUSMOIS(DateDep,i);
            case DelaiMaille[1] of
             '1':code:=DateTimeToStr(DEBUTDEMOIS(DateI));
             '2':code:=DateTimeToStr(DEBUTDEMOIS(DateI)+14);
             '3':code:=DateTimeToStr(FINDEMOIS(DateI));
            end;
           end;
     '6' : begin
            //mois 4-4-5
            Per:=AddPeriodeP(DateDep,i);
            code:=DateTimeToStr(Per.datedeb);
           end;  
     '7' : begin
            //trimestre
            DateI:=PLUSMOIS(DateDep,i*3);
            code:=DateTimeToStr(DEBUTDEMOIS(DateI));
           end;
     '8' : begin
            //quadrimestre
            DateI:=PLUSMOIS(DateDep,i*4);
            code:=DateTimeToStr(DEBUTDEMOIS(DateI));
           end;
     else code:='';
    end;
   end
  else
  //-----------------> ORLI
   begin
    code:='M'+IntTostr(i+1);
   end;
  //ORLI <-----------------
 result:=code;
end;

//Controle si les deux dates données sont renseignées (<>0)
//et que date Début < date Fin
function ControleDateIntervalle(DateDeb,DateFin:TDateTime):boolean;
begin
 result:=true;
 if (DateDeb=0) or (DateFin=0) or (DateDeb>DateFin)
  then result:=false;
end;

//retourne vrai si la date donnée
//appartient bien à l'intervalle date début, date de fin
function DateAppartIntervalle(DateD,DateDeb,DateFin:TDateTime):boolean;
begin
 result:=false;
 if (DateD>=DateDeb) and (DateD<=DateFin)
  then result:=true;
end;

procedure DonneDateDebEtFinSaison(const Saison:hString;
          var dateDeb,dateFin:TDateTime);
var Q:TQuery;
begin
 dateDeb:=0;
 dateFin:=0;
 Q:=MOpenSql('SELECT QSA_SAISONDEBUT,QSA_SAISONFIN FROM QSAISON '+
             'WHERE '+whereCtx(['qsaison'])+' AND QSA_SAISON="'+Saison+'"',
             'DonneDateDebEtFinSaison (BPBasic).',true);
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
             'BPBasic (ChercheDateDDateFPeriode).',true);
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

{ EVI / Modif pour respect semaine 53 }
function NbSemaineIntervalle(DateDeb,DateFin:TDateTime):integer;
var NumSemDeb,NumSemFin,anDeb,anFin:integer;
anIntermediaire,NbSemAnDeb,NbSemAnFin,NbSemAnInt : integer;

function NbSemaine (Annee : integer) : integer;
var  NbSemAn : integer;
DateFinAn : TDateTime;
begin
  NbSemAn :=1 ;
  DateFinAn := StrToDateTime('31/12/'+IntToStr(Annee));
  While NbSemAn = 1 do
  begin
    NbSemAn := NumSemaine(DateFinAn,anDeb);
    DateFinAn := DateFinAn -1;
  end;
result := NbSemAn;
end;

begin
  NbSemAnInt := 0;

  NumSemDeb:=NumSemaine(DateDeb,anDeb);
  NumSemFin:=NumSemaine(DateFin,anFin);

  if anDeb=anFin then result:=NumSemFin-NumSemDeb+1
  else
  begin
    NbSemAnDeb := NbSemaine(anDeb) - NumSemDeb + 1;
    NbSemAnFin := NumSemFin;

    for anIntermediaire:= (anDeb+1)  to (anFin-1)
    do
    begin
      NbSemAnInt := NbSemAnInt + NbSemaine(AnIntermediaire)
    end;

    result := NbSemAnDeb + NbSemAnInt + NbSemAnFin;
  end;
end;

function NbMoisIntervalle(DateDeb,DateFin:TDateTime):integer;
var anDeb,moisDeb,jourDeb,anF,moisF,jourF:word;
begin
 result:=0;
 DecodeDate(DateDeb,anDeb,moisDeb,jourDeb);
 DecodeDate(DateFin,anF,moisF,jourF);
 if anDeb=anF
  then result:=moisF-MoisDeb+1;
 if anF>anDeb
  then result:=(12-MoisDeb)+moisF+1;
end;

function QuestBPInit(const codeSession:hString):boolean;
var rep:integer;
begin
 result:=false;
 if ExisteSql('SELECT QBR_CODESESSION FROM QBPARBRE '+
              'WHERE QBR_CODESESSION="'+codeSession+'"')
  then
   begin
    rep:=HShowmessage('1;Initialisation;Des données existent pour cette session, elles vont être écrasées.'+#13#10+' Etes-vous sûr ?;Q;YN;N;N', '', '');
    if rep=MrYes
     then result:=true;
   end
  else result:=true;
end;

function QuestBPCalend(const codeSession:hString):boolean;
var rep:integer;
begin
 result:=false;
 if ExisteSql('SELECT QBE_CALENDREP FROM QBPDETCALENDREP '+
              'WHERE QBE_CALENDREP="'+codeSession+'"')
  then
   begin
    rep:=HShowmessage('1;Initialisation;Des calendriers existent pour cette session, ils vont être écrasés.'+#13#10+' Etes-vous sûr ?;Q;YN;N;N', '', '');
    if rep=MrYes then
    begin
      result:=true;
      ExecuteSQL('DELETE FROM QBPDETCALENDREP '+
              'WHERE QBE_CALENDREP="'+codeSession+'"')
    end;
   end
  else result:=true;
end;


function QuestionBP():boolean;
var rep:integer;
begin
 result:=false;
 rep:=HShowmessage('1;Initialisation;Des données existent pour cette session, elles vont être écrasées.'+#13#10+' Etes-vous sûr ?;Q;YN;N;N', '', '');
 if rep=MrYes
  then result:=true;
end;

//retourne le mode d'initialisation mois/semaine/quinzaine
function SessionBPInitialise(const session:hString):hString;
var Q:TQuery;
begin
 Q:=MOPenSql('SELECT QBS_BPINITIALISE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+session+'" ',
             'BPBasic (SessionBPInitialise).',true);
 if not Q.eof
  then result:=Q.fields[0].asString;
 ferme(Q);
end;

//*****************************************************************************
//

function NbTailleDansGrilleTaille(GrilleTaille:hString;
         var TabTai:array of hString):integer;
var Q:TQuery;
    i:integer;
begin
 for i:=1 to 20 do
  TabTai[i]:='';  
 i:=0;
 Q:=MOpenSql('SELECT QTA_TAILLE1,QTA_TAILLE2,QTA_TAILLE3,QTA_TAILLE4,QTA_TAILLE5,'+
             'QTA_TAILLE6,QTA_TAILLE7,QTA_TAILLE8,QTA_TAILLE9,QTA_TAILLE10,'+
             'QTA_TAILLE11,QTA_TAILLE12,QTA_TAILLE13,QTA_TAILLE14,QTA_TAILLE15,'+
             'QTA_TAILLE16,QTA_TAILLE17,QTA_TAILLE18,QTA_TAILLE19,QTA_TAILLE20 '+
             ' FROM QTAILLE WHERE '+WhereCtx(['qtaille']) +
             ' AND QTA_CODETAILLE="'+GrilleTaille+'"',
             'BPBasic (NbTailleDansGrilleTaille).', true);
 if not Q.eof
  then
   begin
    while (i<20) and (Q.fields[i].asString<>'') do
     begin
      TabTai[i]:=Q.fields[i].asString;
      i:=i+1;
     end;
   end;
 ferme(Q);
 result:=i;
end;

//retourne le libellé d'un codeaxe donné
function DonneLibelleAxe(codeaxe:hString):hString;
var Q:TQuery;
  CodePremier,CodeTablette,CodeTypeLibre,sLibelle : String;
begin
  result:='';

  {$IFDEF PAIEGRH}
  if CodeMarcheCommun='QAP' then
  begin
    if codeaxe='002' then result:=VH_Paie.PgLibelleOrgStat1;
    if codeaxe='003' then result:=VH_Paie.PgLibelleOrgStat2;
    if codeaxe='004' then result:=VH_Paie.PgLibelleOrgStat3;
    if codeaxe='005' then result:=VH_Paie.PgLibelleOrgStat4;
    if codeaxe='006' then result:=VH_Paie.PgLibCombo1;
    if codeaxe='007' then result:=VH_Paie.PgLibCombo2;
    if codeaxe='008' then result:=VH_Paie.PgLibCombo3;
    if codeaxe='009' then result:=VH_Paie.PgLibCombo4;
  end;

  if result = '' then
  begin
  {$ENDIF PAIEGRH}
    Q:=MOPenSql('SELECT CO_LIBELLE FROM COMMUN WHERE '+
                'CO_CODE="'+codeaxe+'" AND CO_TYPE="'+CodeMarcheCommun+'"',
                'BPBasic (DonneLibelleAxe).',true);
    if not Q.eof then
    Begin
      result:=Q.fields[0].AsString;
      // GC_20080331_GM_FQ_GC15917 DEBUT
      if CodeMarcheCommun='QAE' then
      Begin
        sLibelle := Q.fields[0].AsString;
        CodePremier := ReadTokenPipe(sLibelle,';');
        CodeTablette := ReadTokenPipe(sLibelle,';');
        CodeTypeLibre := ReadTokenPipe (sLibelle,';');
        If (CodePremier = '&#@') Then
          Result :=  RechDomZoneLibre(CodeTypeLibre,False);
      End;
    End;
    // GC_20080331_GM_FQ_GC15917 FIN
    ferme(Q)
  {$IFDEF PAIEGRH}
  end;
  {$ENDIF PAIEGRH}

end;

//donne le libellé d'un code axe
//suivant le code structure et le numéro du code axe
function DonneLibelleCodeAxe(const codestruct,num:hString):hString;
var Q:TQuery;
  CodePremier,CodeTablette,CodeTypeLibre,sLibelle : String;
begin
 result:='';
 if num<>''
  then
   begin
    result:='Axe'+num;

      Q:=MOPenSql('SELECT CO_LIBELLE FROM COMMUN WHERE '+
                  'CO_CODE=(SELECT QBC_CODEAXES'+num+' FROM QBPSTRUCTURE '+
                  'WHERE QBC_CODESTRUCT="'+codestruct+'") AND CO_TYPE="'+CodeMarcheCommun+'"',
                  'BPBasic (DonneLibelleCodeAxe).',true);

    if not Q.eof Then
    Begin
      Result:=Q.fields[0].AsString;
      // GC_20080331_GM_FQ_GC15917 DEBUT
      if CodeMarcheCommun='QAE' then
      Begin
        sLibelle := Q.fields[0].AsString;
        CodePremier := ReadTokenPipe(sLibelle,';');
        CodeTablette := ReadTokenPipe(sLibelle,';');
        CodeTypeLibre := ReadTokenPipe (sLibelle,';');
        If (CodePremier = '&#@') Then
          Result :=  RechDomZoneLibre(CodeTypeLibre,False);
      End;
    End;
    // GC_20080331_GM_FQ_GC15917 FIN
    ferme(Q);
   end;
end;

function DonneLibelleCAxe(const codestruct,num,CodeAxe:hString):hString;
begin
 result:='';
 if BPOkOrli
    //-----------------> ORLI
  then result:=UPPERCASE(RemplaceEspaceParTrait(DonneLibelleCodeAxe(codestruct,num)))
    //ORLI <-----------------
  else result:=DonneLibelleAxe(CodeAxe);
end;


//retourne vrai si le codeaxe entré existe dans la table biblioaxe
//et si c'est un code axe PGI
function ControleAxePGI(const codeAxe,msg:hString):boolean;
begin
  result:=false;
  if codeAxe='' then
  begin
    result:=true;
    exit;
  end;

  if ExisteSql('SELECT QBX_CODEAXE FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+codeaxe+
              '" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"')
  then result:=true
  else PGIINFO(Format_(
         traduireMemoire('La valeur %s du code axe %s est incorrecte.'),
               [codeAxe,msg]));

end;

//retourne vrai si le codeaxe entré existe dans la table biblioaxe
//et si c'est un code axe non PGI
function ControleAxeAutre(const codeAxe,structure,msg:hString):boolean;
begin
 result:=false;
 if codeAxe=''
  then
   begin
    result:=true;
    exit;
   end;
 if ExisteSql('SELECT QBC_CODESTRUCT FROM QBPSTRUCTURE '+
              'WHERE QBC_CODESTRUCT="'+structure+'" AND ('+
              'QBC_CODEAXES1="'+codeaxe+'" OR QBC_CODEAXES2="'+codeaxe+'" OR '+
              'QBC_CODEAXES3="'+codeaxe+'" OR QBC_CODEAXES4="'+codeaxe+'" OR '+
              'QBC_CODEAXES5="'+codeaxe+'" OR QBC_CODEAXES6="'+codeaxe+'" OR '+
              'QBC_CODEAXES7="'+codeaxe+'" OR QBC_CODEAXES8="'+codeaxe+'" OR '+
              'QBC_CODEAXES9="'+codeaxe+'" OR QBC_CODEAXES10="'+codeaxe+'" OR '+
              'QBC_CODEAXES11="'+codeaxe+'" OR QBC_CODEAXES12="'+codeaxe+'" OR '+
              'QBC_CODEAXES13="'+codeaxe+'" OR QBC_CODEAXES14="'+codeaxe+'" OR '+
              'QBC_CODEAXES15="'+codeaxe+'" OR QBC_CODEAXES16="'+codeaxe+'" OR '+
              'QBC_CODEAXES17="'+codeaxe+'" OR QBC_CODEAXES18="'+codeaxe+'" OR '+
              'QBC_CODEAXES19="'+codeaxe+'" OR QBC_CODEAXES20="'+codeaxe+'" OR '+
              'QBC_CODEAXES21="'+codeaxe+'" OR QBC_CODEAXES22="'+codeaxe+'" OR '+
              'QBC_CODEAXES23="'+codeaxe+'" OR QBC_CODEAXES24="'+codeaxe+'" OR '+
              'QBC_CODEAXES25="'+codeaxe+'" OR QBC_CODEAXES26="'+codeaxe+'" OR '+
              'QBC_CODEAXES27="'+codeaxe+'" OR QBC_CODEAXES28="'+codeaxe+'" OR '+
              'QBC_CODEAXES29="'+codeaxe+'" OR QBC_CODEAXES30="'+codeaxe+'" )')
  then result:=true
  else PGIINFO(Format_(
      traduireMemoire('La valeur %s du code axe %s est incorrecte pour le code structure %s.'),
      [codeAxe, msg, structure]));
end;


//retourne vrai si le codeaxe est correct
//cas général
function ControleAxe(const codeAxe,structure,msg:hString):boolean;
begin
 if BPOkOrli
  then result:=ControleAxeAutre(codeAxe,Structure,msg)
  else result:=ControleAxePGI(codeAxe,msg);
end;

//retourne le nom de la table, le champ code , le champ libelle
//suivant code axe
procedure DonneNomTableChampsPrValeurAxe(const CodeAxe: hString;
          var NomTable,ChpCode,ChpLib,Chp2,Val2:hString);
var Q:TQuery;
begin
  Chp2:='';
  Val2:='';
  Q:=MOpenSql('SELECT QBX_NOMTABLEAXE,QBX_CHPCODE,QBX_CHPLIB,'+
             'QBX_NOMCHP2,QBX_VALCHP2 '+
             'FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="'+codeaxe+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
             'BPBasic (DonneNomTableChampsPrValeurAxe).',true);
  if not Q.eof
   then
    begin
     NomTable:=Q.fields[0].asString;
     ChpCode:=Q.fields[1].asString;
     ChpLib:=Q.fields[2].asString;
     Chp2:=Q.fields[3].asString;
     Val2:=Q.fields[4].asString;
    end;
   ferme(Q);
 if NomTable=''
  then
   begin
    NomTable:='QBPAXE';
    ChpCode:='QBA_VALEURAXE';
    ChpLib:='QBA_LIBVALEURAXE';
    Chp2:='QBA_CODEAXE';
    Val2:=CodeAxe;
   end;
end;

//retourne vrai si la valeur d'axe entrée est correcte
//par rapport au code axe donné
//cas non PGI
function ControleValeurAxeAutre(const codeAxe,valeurAxe,msg:hString):boolean;
begin
 result:=false;
 if codeAxe=''
  then
   begin
    result:=true;
    exit;
   end;

 if valeurAxe=''
  then
   begin
    result:=false;
    PGIINFO(Format_(
         traduireMemoire('Vous devez renseigner une valeur d''axe pour le code axe %s (%s).'),
               [codeAxe,msg]));
    exit;
   end;

 if ExisteSql('SELECT QBA_CODEAXE FROM QBPAXE '+
              'WHERE QBA_CODEAXE="'+codeaxe+
              '" AND QBA_VALEURAXE="'+valeurAxe+'"')
  then result:=true
  else PGIINFO(Format_(
         traduireMemoire('La valeur %s est incorrecte pour le code axe %s (%s).'),
               [valeurAxe,codeAxe,msg]));
end;

//retourne vrai si la valeur d'axe entrée est correcte
//par rapport au code axe donné
//cas PGI
function ControleValeurAxePGI(const codeAxe,valeurAxe,msg:hString):boolean;
var NomTable,ChpCode,ChpLib,Chp2,Val2,codeS:hString;
begin
 result:=false;
 if codeAxe=''
  then
   begin
    result:=true;
    exit;
   end;
 
 if valeurAxe=''
  then
   begin
    result:=false;
    PGIINFO(Format_(
         traduireMemoire('Vous devez renseigner une valeur d''axe pour le code axe %s (%s).'),
               [codeAxe,msg]));
    exit;
   end;

 DonneNomTableChampsPrValeurAxe(CodeAxe,NomTable,ChpCode,ChpLib,Chp2,Val2);

 codeS:='';
 if chp2<>''
  then codeS:=' AND '+chp2+'="'+val2+'" ';
 if ExisteSql('SELECT '+ChpCode+' FROM '+NomTable+' WHERE '+chpcode+'="'+valeuraxe+'" '+codeS)
  then result:=true
  else PGIINFO(Format_(
         traduireMemoire('La valeur %s est incorrecte pour le code axe %s (%s).'),
               [valeurAxe,codeAxe,msg]));
end;

//retourne vrai si la valeur d'axe entrée est correcte
//par rapport au code axe donné
//cas général
function ControleValeurAxe(const codeAxe,valeurAxe,msg:hString):boolean;
var S,D:hString;
    ii,i,L:integer;
    ResOk:boolean;
begin
 result:=true;
 ResOk:=true;
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
   if BPOkOrli
      //-----------------> ORLI
    then ResOk:=ResOk and ControleValeurAxeAutre(codeAxe,D,msg)
      //ORLI <-----------------
    else ResOk:=ResOk and ControleValeurAxePGI(codeAxe,D,msg);
   ii:=ii+1;
  end;

 if ii=0
  then
   begin
    if BPOkOrli
      //-----------------> ORLI
     then ResOk:=ResOk and ControleValeurAxeAutre(codeAxe,S,msg)
      //ORLI <-----------------
     else ResOk:=ResOk and ControleValeurAxePGI(codeAxe,S,msg);
   end
  else
   begin
    if S<>''
     then
     if BPOkOrli
        //-----------------> ORLI
      then ResOk:=ResOk and ControleValeurAxeAutre(codeAxe,S,msg)
        //ORLI <-----------------
      else ResOk:=ResOk and ControleValeurAxePGI(codeAxe,S,msg);
   end;
 result:=ResOk;
end;

function ChercheMaxCompteurPivot:integer;
var Q:TQuery;
    X:hString;
begin
 result:=0;
 Q:=MOpenSql('SELECT MAX(CAST(QBV_CODEIDENT AS FLOAT)) FROM QBPPIVOT ',
             'BPBasic (ChercheMaxCompteurPivot).',true);
 if not Q.eof
  then X:=(Q.fields[0].asString);
 ferme(Q);
 if X<>''  
  then result:=VALEURI(X);
end;

function ObjPrevConsultant():boolean;
var Q:TQuery;
begin
 result:=false;
 if BPOkOrli
  then
  //-----------------> ORLI
   begin
    Q:=MOpenSql('SELECT QEV_CONSULTANT FROM QENVIRONNEMENT '+
                'WHERE QEV_ENVIRONNEMENT="'+UserDonneEnvironnement+'"',
                'BPBasic (ObjPrevConsultant).', true);
    if not Q.eof then
     begin
       if Q.fields[0].asString='X'
        then result:=true;
     end;
    ferme(Q);
   end; 
  //ORLI <-----------------
end;



//retourne le code pour la tablette
//pour avoir les différents codes axes qui existent dans la structure donnée
function DonneCodeTabletteAxeStructure(const structure:hString):hString;
var Q:TQuery;
    i:integer;
    res,code:hString;
begin
 res:='';
 code:='';
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
             ' FROM QBPSTRUCTURE WHERE QBC_CODESTRUCT="'+structure+'"',
             'BPBasic (DonneCodeTabletteAxeStructure).',true);
 if not Q.eof
  then
   begin
    for i:=1 to 30 do
     begin
      code:=Q.fields[i].asString;
      if code<>''
       then res:=res+' OR QBX_CODEAXE="'+code+'" '
     end;
   end;
 ferme(Q);
 result:=res;
end;

function ChercheTabletteValeurAxe(const codeAxe:hString):hString;
var Q:TQuery;
begin
 result:='';

 Q:=MOPenSql('SELECT QBX_NOMTABLETTE FROM QBPBIBLIOAXE '+
             'WHERE QBX_CODEAXE="'+codeAxe+'" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"',
             'BPBasic (ChercheTabletteValeurAxe).',true);
 if not Q.eof
  then result:=Q.fields[0].asString;
 ferme(Q);
end;

//retourne vrai si il y a une loi de definie pour la session donnée
//faux sinon
function  OkSessionLoiDefinie(const CodeSession:hString):boolean;
begin
 result:=false;
 if ExisteSql('SELECT QBO_CODESESSION FROM QBPLOI '+
              'WHERE QBO_CODESESSION="'+CodeSession+'"')
  then result:=true;

end;


procedure ValRestrictionValaxe1234(valRestriction:hString;
          var valaxe1,valaxe2,valaxe3,valaxe4:hString);
var S:hString;
    i1,i2,i3,i4,L:integer;
begin
 S:=valRestriction;
 L:=length(S);
 i1:=pos('champ1=',S);
 i2:=pos('champ2=',S);
 i3:=pos('champ3=',S);
 i4:=pos('champ4=',S);

 valaxe1:=copy(S,i1+7,i2-7-i1);
 valaxe2:=copy(S,i2+7,i3-7-i2);
 valaxe3:=copy(S,i3+7,i4-7-i3);
 valaxe4:=copy(S,i4+7,L-+1-7-i4);
end;

function Valaxe1234ValRestriction(valaxe1,valaxe2,valaxe3,valaxe4:hString):hString;
begin
 result:='champ1='+valaxe1+'champ2='+valaxe2+'champ3='+valaxe3+'champ4='+valaxe4;
end;


function DonneLibelleDelai(Delai:TDateTime;BPInitialise:hString):hString;
var NumSem,An:integer;
    code:hString;
    days: array[1..7] of hString;
    mois: array[1..12] of hString;
    y,m,d:word;
    per: TperiodeP;
begin
 days[1] := TraduireMemoire('Dimanche');
 days[2] := TraduireMemoire('Lundi');
 days[3] := TraduireMemoire('Mardi');
 days[4] := TraduireMemoire('Mercredi');
 days[5] := TraduireMemoire('Jeudi');
 days[6] := TraduireMemoire('Vendredi');
 days[7] := TraduireMemoire('Samedi');
 mois[1] := TraduireMemoire('Janvier');
 mois[2] := TraduireMemoire('Février');
 mois[3] := TraduireMemoire('Mars');
 mois[4] := TraduireMemoire('Avril');
 mois[5] := TraduireMemoire('Mai');
 mois[6] := TraduireMemoire('Juin');
 mois[7] := TraduireMemoire('Juillet');
 mois[8] := TraduireMemoire('Août');
 mois[9] := TraduireMemoire('Septembre');
 mois[10] := TraduireMemoire('Octobre');
 mois[11] := TraduireMemoire('Novembre');
 mois[12] := TraduireMemoire('Décembre');
 case BPInitialise[1] of
  '1':code:=days[DayOfWeek(Delai)];
  '2':begin
       NumSem:=NumSemaine(Delai,An);
       code:='Sem '+IntToStr(NumSem)+'-'+IntToStr(An);
      end;
  '3':begin
       decodeDate(Delai,y,m,d);    
       case d of
        1:code:=TraduireMemoire('1° quinzaine ')+MetZero(intTostr(m),2)+'-'+intTostr(y);
        16:code:=TraduireMemoire('2° quinzaine ')+MetZero(intTostr(m),2)+'-'+intTostr(y);
       end;

      end;
  '4':begin
       decodeDate(Delai,y,m,d);
       code:=mois[m]+' '+IntToStr(y);
      end;  
  '5':begin
       Per:=FindPeriodeP(Delai);
       code:=Per.libelle;
      end;
  '6':begin
       decodeDate(Delai,y,m,d);
       case m of
        1:code:=TraduireMemoire('1° trimestre ')+intTostr(y);
        4:code:=TraduireMemoire('2° trimestre ')+intTostr(y);
        7:code:=TraduireMemoire('3° trimestre ')+intTostr(y);
        10:code:=TraduireMemoire('4° trimestre ')+intTostr(y);
       end;
      end;
  '7':begin
       decodeDate(Delai,y,m,d);
       case m of
        1:code:=TraduireMemoire('1° quadrimestre ')+intTostr(y);
        5:code:=TraduireMemoire('2° quadrimestre ')+intTostr(y);
        9:code:=TraduireMemoire('3° quadrimestre ')+intTostr(y);
       end;
      end;
  else code:=DateTimeToStr(delai);
 end;
 result:=code;
end;

//procedure qui donne
//pour le code session donné
function ChercheNumOrdreSession(const codeSession:hString):hString;
var Q:TQuery;
    sessionEclat,sessionEclatTai:hString;
begin
 result:='GLOBAL';
 Q:=MOPenSql('SELECT QBS_SESSIONECLAT,QBS_SESSIONECLATAI '+
             'FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'BPBasic (ChercheNumOrdreSession).',true);
 if not Q.eof
  then
   begin
    sessionEclat:=Q.fields[0].asString;
    sessionEclatTai:=Q.fields[1].asString;
    if (sessionEclat='X')
     then result:='DELAI';
   end;
 ferme(Q);
end;

//pour une structure donnée
//donne le nivMax càd le nombre de codes axes renseignés
//et le tableau des axes
procedure ChercheAxeStructure(const structure:hString;
          var TabCodeAxe:array of hString;
          var NivMax:integer);
var Q:TQuery;
    i:integer;
    code:hString;
begin
  NivMax := 0;
  for i := 1 to 30 do
    TabCodeAxe[i] := '';
  if structure = '' then
    exit;
  Q := MOpenSql('SELECT QBC_CODEAXES1,QBC_CODEAXES2,' +
             'QBC_CODEAXES3,QBC_CODEAXES4,QBC_CODEAXES5,QBC_CODEAXES6,'+
             'QBC_CODEAXES7,QBC_CODEAXES8,QBC_CODEAXES9,QBC_CODEAXES10,'+
             'QBC_CODEAXES11,QBC_CODEAXES12,QBC_CODEAXES13,QBC_CODEAXES14,'+
             'QBC_CODEAXES15,QBC_CODEAXES16,QBC_CODEAXES17,QBC_CODEAXES18,'+
             'QBC_CODEAXES19,QBC_CODEAXES20,QBC_CODEAXES21,QBC_CODEAXES22,'+
             'QBC_CODEAXES23,QBC_CODEAXES24,QBC_CODEAXES25,QBC_CODEAXES26,'+
             'QBC_CODEAXES27,QBC_CODEAXES28,QBC_CODEAXES29,'+
             'QBC_CODEAXES30 FROM QBPSTRUCTURE WHERE QBC_CODESTRUCT="'+structure+'"',
             'BPBasic (ChercheAxeStructure).',true);
 if not Q.eof
  then
   begin
    for i := 0 to 29 do
     begin
      code := trim(Q.fields[i].asString);
      if code<>''
       then
        begin
        TabCodeAxe[i + 1] := code;
        NivMax := i + 1;
        end;
     end;
   end;
 ferme(Q);
end;

function AnalyseJoinSQL(const Join : hString; var TS, TS2, TS3: TstringList): hString;
var
  i, j, k  : integer;
  Position, s, LeftJoin, Table, Alias, AliasOk : hString;

  function SuppJoin(const Indice : integer): hString;
  var
    i     : integer;
    tts, s, t, a, s1 : hString;
  begin
    Result := '';
    tts := TS[Indice];
    s := ReadTokenSt(tts);
    t := ReadTokenSt(tts);
    a := ReadTokenSt(tts);
    for i := Indice + 1 to TS.Count - 1 do
    begin
      s1 := TS[i];
      s1 := ReadTokenSt(s1);
      if s = s1 then
      begin
        Result := Result + MetZero(intToStr(i), 2) + ';';
        TS3[i] := a;
      end;
    end;
  end;

  function SavAlias(const Indice : integer): hString;
  var
    i                     : integer;
    s, s1, lj, lj1, t, t1 : hString;
  begin
    Result := '';
    s  := TS[Indice];
    lj := ReadTokenSt(s);
    t  := ReadTokenSt(s);
    for i := Indice + 1 to TS.Count - 1 do
    begin
      if TS[i] <> '' then
      begin
        s1  := TS[i];
        lj1 := ReadTokenSt(s1);
        t1  := ReadTokenSt(s1);
        if t = t1 then
        begin
          Result := Result + MetZero(intToStr(Indice), 2) + ';';
          Result := Result + MetZero(intToStr(i), 2) + ';';
        end;
      end;
    end;
  end;

  function ChercheProchainEspace( const Chaine : hString): integer;
  var
    i : integer;
  begin
    Result := 0;
    for i:= 1 to length(Chaine)-1 do
    begin
      if (copy(Chaine, i, 1) <> ' ') and (copy(Chaine, i+1, 1) = ' ') then
      begin
        Result := i;
        break;
      end;
    end;
  end;
begin
  Result := '';
  i:= 1;
  s:= Trim(UpperCase(Join));
  //Remplit une TStringList avec les JOIN + ';' Table + ';' + Alias + ';' + Alias à sauvegarder X ou -
  while i <> 0 do
  begin
    i := pos('LEFT JOIN', Copy(s, 1, Length(s)));
    if i <> 0 then
    begin
      s := Copy(s, 1, i-1) + '&&&& &&&&' + Copy(s, i+9, Length(s));
      j := pos('LEFT JOIN', Copy(s, 1, Length(s)));
      if j <> 0 then
        LeftJoin := Copy(s, i, j-i)
      else
        LeftJoin := Copy(s, i, Length(s));
    end;
    if i <> 0 then
    begin
      k := pos('&&&& &&&&', Copy(LeftJoin, 1, Length(LeftJoin)));
      Table := Copy(LeftJoin, k+9, ChercheProchainEspace(Copy(LeftJoin, k+9, Length(LeftJoin))));
      Alias := Copy(LeftJoin, k+9+length(Table), ChercheProchainEspace(Copy(LeftJoin, k+9+length(Table), Length(LeftJoin))));
      TS.Add(Trim(LeftJoin) + ';' + Trim(Table) + ';' + Trim(Alias));
      TS2.Add(Trim(Alias));
      TS3.Add('');
    end;
  end;
  //Suppression des Join redondants + Suppression des Alias
  if assigned(TS) and (TS.Count > 0) then
  begin
    for k := 0 to TS.Count -1 do
    begin
      s := TS[k];
      LeftJoin := ReadTokenSt(s);
      Table    := ReadTokenSt(s);
      Alias    := ReadTokenSt(s);
      LeftJoin := StringReplace(LeftJoin, Alias, '##', [rfReplaceAll]);
      TS[k] := Trim(LeftJoin) + ';' + Trim(Table) + ';' + Trim(Alias);
    end;
    Position := '';
    for k := 0 to TS.Count -1 do
    begin
      if (pos(MetZero(intToStr(k), 2), Position)=0) then
        Position := Position + SuppJoin(k);
    end;
    //RAZ de la TStringList TS pour les Join redondants
    while Position <> '' do
    begin
      k := ValeurI(ReadTokenSt(Position));
      if k >=0 then TS[k] := '';
    end;
    //Alias à sauvegarder => LEFT JOIN <> sur table identique
    Position := '';
    for k := 0 to TS.Count -1 do
    begin
      if (TS[k] <> '') and (pos(MetZero(intToStr(k), 2), Position) = 0) then
        Position := Position + SavAlias(k);
    end;
    for k := 0 to TS.Count -1 do
    begin
      if (pos(MetZero(intToStr(k), 2), Position) > 0) then
      begin
        if TS[k] <> '' then TS[k]  := TS[k]  + ';X';
        TS2[k] := TS2[k] + ';X';
      end
      else
      begin
        if TS[k] <> '' then TS[k]  := TS[k]  + ';-';
        TS2[k] := TS2[k] + ';-';
      end;
    end;
    //Suppression de la TStringList des Alias (sauf si LEFT JOIN <> sur table identique)
    //et restitution mot clé LEFT JOIN
    for k := 0 to TS.Count -1 do
    begin
      s := TS[k];
      if s <> '' then
      begin
        LeftJoin := ReadTokenSt(s);
        Table    := ReadTokenSt(s);
        Alias    := ReadTokenSt(s);
        AliasOk  := ReadTokenSt(s);
        TS[k] := StringReplace(TS[k], '&&&& &&&&', 'LEFT JOIN', [rfReplaceAll]);
        if not StrToBool_(AliasOk) then
        begin
          TS[k] := StringReplace(TS[k], '##.', '', [rfReplaceAll]);
          TS[k] := StringReplace(TS[k], '##' , '', [rfReplaceAll]);
        end
        else
        begin
          TS[k] := StringReplace(TS[k], '##.', Alias+'.', [rfReplaceAll]);
          TS[k] := StringReplace(TS[k], '##' , Alias    , [rfReplaceAll]);
        end;

        s := TS[k];
        Result := Result + ' ' + ReadTokenSt(s);
      end;
    end;
  end;
end;

function AnalyseJoinSQLPaie(const Join,TablePrincipale : hString): hString;
var
  i, j, k  : integer;
  s, LeftJoin, Table : hString;

  function ChercheProchainEspace( const Chaine : hString): integer;
  var
    i : integer;
  begin
    Result := 0;
    for i:= 1 to length(Chaine)-1 do
    begin
      if (copy(Chaine, i, 1) <> ' ') and (copy(Chaine, i+1, 1) = ' ') then
      begin
        Result := i;
        break;
      end;
    end;
  end;
begin
  Result := '';
  i:= 1;
  s:= Trim(UpperCase(Join));
  //Remplit une TStringList avec les JOIN + ';' Table + ';' + Alias + ';' + Alias à sauvegarder X ou -
  while i <> 0 do
  begin
    i := pos('LEFT JOIN', Copy(s, 1, Length(s)));
    if i <> 0 then
    begin
      s := Copy(s, 1, i-1) + '&&&& &&&&' + Copy(s, i+9, Length(s));
      j := pos('LEFT JOIN', Copy(s, 1, Length(s)));
      if j <> 0 then
        LeftJoin := Copy(s, i, j-i)
      else
        LeftJoin := Copy(s, i, Length(s));
    end;
    if i <> 0 then
    begin
      k := pos('&&&& &&&&', Copy(LeftJoin, 1, Length(LeftJoin)));
      Table := Copy(LeftJoin, k+9, ChercheProchainEspace(Copy(LeftJoin, k+9, Length(LeftJoin))));
    end;
  end;
  if Trim(UpperCase(TablePrincipale)) = Trim(UpperCase(Table)) then Result := ''
  else Result := Join;
end;


function BPTablettePlus(codeStruct: hString): hString;
begin
  result := '';
  Case ContextBP of
    0,2,3 : result:='QBX_CODEMARCHE="#PGIENT"';
    1 : result:='QBX_CODEMARCHE="#PGI"';
  end;
 if BPOkOrli
    //-----------------> ORLI
  then result:='QBX_CODEMARCHE<>"#PGI" and QBX_CODEMARCHE<>"#PGIENT"';
    //ORLI <-----------------
end;

function BPLibelleMaille(BPinitialise:hString):hString;
var code:hString;
begin
 code:='';
 if BpInitialise<>''
  then
 case BPInitialise[1] of
  '1':code:=TraduireMemoire('au jour');
  '2':code:=TraduireMemoire('à la semaine');
  '3':code:=TraduireMemoire('à la quinzaine');
  '4':code:=TraduireMemoire('au mois');
  '5':code:=TraduireMemoire('au mois 4-4-5');
  '6':code:=TraduireMemoire('au trimestre');
  '7':code:=TraduireMemoire('au quadrimestre');
 else code:=TraduireMemoire(' au global');
 end;

 result:=code;
end;

//Renvoit vrai si l'axe Devise est sélectioné
function AxeDeviseOK(const session: hString): boolean;
var Q: TQuery;
i,NivMax : integer;
begin
  NivMax := ChercheNivMaxSession(session);

  Result:=false;

  Q:=MopenSql('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4'+
              ',QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,'+
              'QBS_CODEAXES9,QBS_CODEAXES10 '+
              'FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
              'BPFctSession (AxeDeviseOK).',true);
  if not Q.eof then
  begin
    for i:= 0 to NivMax-1 do
    begin
      if Q.fields[i].asString = '011' then
       begin
         result := true;
         break;
       end
    end;
  end;
  ferme(Q);

end;

//retourne le délai minimum des lois définit pour la session en Loi D'Eclatement
function SessionMinLoi(const session:hString):hString;
var Q:TQuery;
begin

 Q:=MOPenSql('SELECT MIN(QBO_MAILLE) FROM QBPLOI '+
             'WHERE QBO_CODESESSION="'+session+'" ',
             'BPBasic (SessionBPInitialise).',true);
 if not Q.eof then result:=Q.fields[0].asString;
 ferme(Q);
 if result='' then result := '1';
end;

function CodeMarcheBiblioAxe : string;
begin
  Case contextBP of
    0 : result:='#PGIENT';
    1 : result:='#PGI';
    2 : result:='#PGICOMPTA';
    3 : result:='#PGIPAIE';
  end;
end;

function CodeMarcheCommun : string;
begin
  Case contextBP of
    0 : result:='QAE';
    1 : result:='QAM';
    2 : result:='QAC';
    3 : result:='QAP';
  end;
end;

function VerifDateCube(Date : hString) : hString;
var a, m, j : Word;
Erreur:boolean;
begin
  a:=0; m:=0; j:=0;
  Erreur := false;

  try
    StrToDate(Date);
    Result := Date;
  except
    on E: EConvertError do
    begin
      if Pos(TraduireMemoire('JANVIER'),UPPERCASE(Date)) > 0 then m := 1;
      if Pos(TraduireMemoire('FEVRIER'),UPPERCASE(Date)) > 0 then m := 2;
      if Pos(TraduireMemoire('MARS'),UPPERCASE(Date)) > 0 then m := 3;
      if Pos(TraduireMemoire('AVRIL'),UPPERCASE(Date)) > 0 then m := 4;
      if Pos(TraduireMemoire('MAI'),UPPERCASE(Date)) > 0 then m := 5;
      if Pos(TraduireMemoire('JUIN'),UPPERCASE(Date)) > 0 then m := 6;
      if Pos(TraduireMemoire('JUILLET'),UPPERCASE(Date)) > 0 then m := 7;
      if Pos(TraduireMemoire('AOUT'),UPPERCASE(Date)) > 0 then m := 8;
      if Pos(TraduireMemoire('SEPTEMBRE'),UPPERCASE(Date)) > 0 then m := 9;
      if Pos(TraduireMemoire('OCTOBRE'),UPPERCASE(Date)) > 0 then m := 10;
      if Pos(TraduireMemoire('NOVEMBRE'),UPPERCASE(Date)) > 0 then m := 11;
      if Pos(TraduireMemoire('DECEMBRE'),UPPERCASE(Date)) > 0 then m := 12;

      try j := StrToInt(Copy(Date,0,2));
      except on E: EConvertError do Erreur := true;
      end;

      if Erreur = false then
      begin
        try a := StrToInt(Copy(Date,length(Date)-3,4));
        except on E: EConvertError do Erreur := true;
        end;

        if Erreur = false then
        begin
          if IsValidDate(a,m,j) then
          begin
            try Result := DateTimeToStr(EncodeDate(a,m,j));
            except on E: EConvertError do Result := '';
            end;
          end else Result := '';
        end else Result := '';
      end else Result := '';
    end;
  end;
end;

end.
