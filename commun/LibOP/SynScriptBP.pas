unit SynScriptBP;

interface
uses {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
      {$IFNDEF EAGLSERVER}Fe_Main,{$ENDIF !EAGLSERVER}
     {$ELSE}
      MainEagl,
      utob,
     {$ENDIF}
     {$IFDEF BPORLI}FichValeurAxe,{$ENDIF}
      HCtrls,controls;

function FuncSynBPBtnInitDelaiObj(parms: array of variant; nb: integer): variant;
function FuncSynBPLanceCalculInitObjectif(parms: array of variant; nb: integer):variant;
procedure ProcSynBPLanceCalculEclatementLoi(parms: array of variant; nb: integer);
function FuncSynBPLibelleMaille(parms: array of variant; nb: integer): variant;
function FuncSynBPSessionValide(parms: array of variant; nb: integer): variant;
function FuncSynBPValeurUnique(parms: array of variant; nb: integer): variant;
function FuncSynBPTabletteCodeAxe(parms: array of variant; nb: integer): variant;
function FuncSynMakeDate(parms: array of variant; nb: integer): variant;
function FuncSynBPDateDebutAnneeMoinsUn(parms: array of variant; nb: integer): variant;
function FuncSynBPDateFinAnneeMoinsUn(parms: array of variant; nb: integer): variant;
function FuncSynBPDateFinAnnee(parms: array of variant; nb: integer): variant;
function FuncSynBPDateDebutAnnee(parms: array of variant; nb: integer): variant;
procedure ProcSynBPMAJDatesLoi(parms: array of variant; nb: integer);
function FuncSynBPTabletteValeurAxeC(parms: array of variant; nb: integer): variant;
function FuncSynBPPlusValeurAxe(parms: array of variant; nb: integer): variant ;
procedure ProcSynBPLanceValidation(parms: array of variant; nb: integer);
function FuncSynBPLanceChxValeurAxe(parms: array of variant; nb: integer): variant;
procedure ProcSynBPLanceDeValidation(parms: array of variant; nb: integer);
function NbDeMaille(session:hString;BPinit:hString='InitOK';DateDeb:TDateTime=0;DateFin:TDateTime=0):integer;
function VerifPos( Sens : boolean; Req, Letter : hString; Pos : integer ) : integer;
function NbMajCoeff(codesession:hString):integer;
procedure ValidationBP(const codeSession:hString);
function FuncSynBPSessionDelete(parms: array of variant; nb: integer): variant; // GC_20080124_GM_GC15727
implementation


uses  uutil,M3FP,HMsgBox,Sysutils,HEnt1,
      InitialisationPrev,
      BPEclatement,BPPrevCmd,
      UCtx,CstCommun,BPFctSEssion,BPInitPrev,BPFctArbre,
      BPBasic,BPUtil
//      {$IFDEF CRM}
      {$IF (Defined(GCGC) AND Not Defined(MODE)) OR Defined(CRM)}
      ,ParamSoc
//      {$ENDIF}
      {$IFEND (Defined(GCGC) AND Not Defined(MODE)) OR Defined(CRM)}
      ;

  (* Messages *)
const  TexteMessage : array[ 1..14 ] of hString = (
    {1} 'Pour une session initialisée par semaine, la date sélectionnée doit correspondre au premier jour de la semaine définie en paramètre (%s).',
    {2} 'Voulez-vous corriger la date sélectionnée au %s.',
    {3} 'Pour une session initialisée par quinzaine, la date sélectionnée doit correspondre au premier jour ou au 16 du mois.',
    {4} 'Confirmez-vous la suppression de cet enregistrement ?',
    {5} '(Session : %s).%s Etes-vous sûr ?',
    {6} 'Pour le calcul d''initialisation global :%s Le nombre de jours pour la période d''objectif (%s) doit être égal à celui de la période de référence (%s).',
    {7} 'Pour le calcul d''initialisation global :%s Le nombre de jours pour la période objectif et la période de référence doit être inférieur à 2 ans.',
    {8} 'Sélection : Période de l''objectif = %s jours - Période de référence = %s jours',
    {9} 'Session %s',
    {10}'Duplication avec les lois ?',
    {11}'Pour une session initialisée par mois, la date sélectionnée doit correspondre au premier jour du mois.',
    {12}'Voulez-vous corriger la date sélectionnée au %s.',
    {13}'Pour une session initialisée par trimestre, la date sélectionnée doit correspondre au premier jour d''un trimestre (01/01,01/04,01/07,01/10).',
    {14}'Pour une session initialisée par quadrimestre, la date sélectionnée doit correspondre au premier jour d''un quadrimestre (01/01,01/05,01/09).'
   );


//fonction qui pose une question pour la confirmation de la suppresion
function FuncSynQuestionSupp(parms: array of variant; nb: integer): variant;
begin
  result := HShowMessage( '1;' + hString(parms[0]) + ';' + TraduireMemoire( TexteMessage[ 4 ] ) + ';Q;YN;N;N', '', '' );
end;

function FuncTrouveArgument(parms: array of variant; nb: integer): variant;
begin
  Result:=TrouveArgument(hString(parms[0]),hString(parms[1]),hString(parms[2]));
end;

//function qui renvoie la date du TDateTime entré
function FuncSynDate(parms: array of variant; nb: integer): variant;
var DT: TDateTime;
begin
  DT:= StrToDateTime0(Parms[0]);
  Result:='';
  if DT>10 then
    Result := DateToStr(DT);
end;

//fonction qui renvoie un TdateTime
//suivant un jour et une heure entrés
function FuncSynMakeDate(parms: array of variant; nb: integer): variant;
var JOUR, HEURE: hString;
begin
  if strtodatetime0(parms[0])=0
  then result := 0
  else
   begin
    JOUR := hString(parms[0]);
    HEURE := hString(parms[1]);
    Result := strtodatetime0(JOUR + ' ' + HEURE);
   end;
end;

//*************************************************************************
//
//Fonctions pour la gestion des contextes dans les scripts
//
//*************************************************************************

function FuncSynGetCtx(parms: array of variant; nb: integer): variant;
begin
  Result := ValeurContexte(hString(parms[0]));
end;

function FuncSynStrToFloat(parms: array of variant; nb: integer): variant;
begin
  Result :=VALEUR(hString(parms[0]));
end;

function FuncSynWhereCtx(parms: array of variant; nb: integer): variant;
begin
  Result := WhereCtx([hString(parms[0])]);
end;

//****************************************************************************
//fonction utilisée dans fiche evoultion
//peut être à supprimer

//fonction qui donne la liste des 4 axes lié au profil
function FuncSynBPChercheAxes(parms: array of variant; nb: integer): variant;
var session,profil,axe1,axe2,axe3,axe4,code : hString;
    Q: TQuery;
begin
  result := '';
  session := hString(parms[0]);
  profil:='';
  axe1:='';
  axe2:='';
  axe3:='';
  axe4:='';
  code:='';
  Q := MOpensql('SELECT QBS_CODEPROFIL FROM QBPSESSIONBP '+
       'WHERE QBS_CODESESSION="'+session+'"',
       'SynScriptBP (FuncSynBPChercheAxes).', true);
  if not Q.Eof
   then profil:=Q.fields[0].AsString;
  ferme(Q);
  Q := MOpensql('SELECT QBF_CODEAXE1,QBF_CODEAXE2,QBF_CODEAXE3,QBF_CODEAXE4 '+
       'FROM QBPPROFIL '+
       'WHERE QBF_CODEPROFIL="'+profil+'"',
       'SynScriptBP (FuncSynBPChercheAxes).', true);
  if not Q.Eof
   then
    begin
     axe1:=Q.fields[0].AsString;
     axe2:=Q.fields[1].AsString;
     axe3:=Q.fields[2].AsString;
     axe4:=Q.fields[3].AsString;
    end;
  ferme(Q);
  if axe1<>''
   then code := code +' QBP_CODEAXE="'+axe1+'"';
  if axe2<>''
   then code := code +' OR QBP_CODEAXE="'+axe2+'"';
  if axe3<>''
   then code := code +' OR QBP_CODEAXE="'+axe3+'"';
  if axe4<>''
   then code := code +' OR QBP_CODEAXE="'+axe4+'"';
  result:=code;
end;

//****************************************************************************

//fonction qui controle si axe donné est lié au profil
function FuncSynBPControleAxes(parms: array of variant; nb: integer): variant;
var session,profil,axe1,axe2,axe3,axe4,axe,libelle : hString;
    Q: TQuery;
    res: boolean;
begin
  res := false;
  session := hString(parms[0]);
  axe := hString(parms[1]);
  libelle := hString(parms[2]);
  profil:='';
  axe1:='';
  axe2:='';
  axe3:='';
  axe4:='';
  Q := MOpensql('SELECT QBS_CODEPROFIL FROM QBPSESSIONBP '+
       'WHERE QBS_CODESESSION="'+session+'"',
       'SynScriptBP (FuncSynBPControleAxes).', true);
  if not Q.Eof
   then profil:=Q.fields[0].AsString;
  ferme(Q);
  Q := MOpensql('SELECT QBF_CODEAXE1,QBF_CODEAXE2,QBF_CODEAXE3,QBF_CODEAXE4 '+
       'FROM QBPPROFIL '+
       'WHERE QBF_CODEPROFIL="'+profil+'"',
       'SynScriptBP (FuncSynBPControleAxes).', true);
  if not Q.Eof
   then
    begin
     axe1:=Q.fields[0].AsString;
     axe2:=Q.fields[1].AsString;
     axe3:=Q.fields[2].AsString;
     axe4:=Q.fields[3].AsString;
    end;
  ferme(Q);
  if (axe1=axe) or (axe2=axe) or (axe3=axe) or (axe4=axe)
   then res:=true;
  if res=false
   then
    begin
     HShowMessage('1;Erreur-Contrôle de validité;La valeur de $$ : %% est incorrecte.;W;O;O;O',
          axe, libelle);
    end;
  result:=res;
end;

//****************************************************************************

//fonction qui donne le lib d'un champ code
function FuncSynDonneLibBP(parms: array of variant; nb: integer): variant;
var code, tab, chp, chp1, val1, chp2, val2: hString;
    Q: TQuery;
begin
  result := '';
  tab := hString(parms[0]);
  chp := hString(parms[1]);
  chp1 := hString(parms[2]);
  val1 := hString(parms[3]);
  chp2 := hString(parms[4]);
  val2 := hString(parms[5]);
  code := '';
  if chp1 <> ''
    then code := ' AND ' + chp1 + '="' + val1 + '" ';
  if chp2 <> ''
    then code := code + ' AND ' + chp2 + '="' + val2 + '" ';
  Q := MOpensql('SELECT ' + chp + ' FROM ' + tab + ' WHERE ' +code,
       'SynScriptBP (FuncSynDonneLibBP).', true);
  if not Q.Eof
    then result := Q.fields[0].AsString;
  ferme(Q);
end;

//****************************************************************************

// fonction de test de validité
// test si la valeur du champ entrée appartient bien à la table
//++ non obligatoire
function FuncSynBPControleValiditeCas4(parms: array of variant; nb: integer): variant;
var res: boolean;
  valeur1, valeur2, table, chp1, chp2, libelle, code: hString;
begin
 res := true;
 valeur1 := hString(Parms[0]);
 table := hString(Parms[1]);
 chp1 := hString(Parms[2]);
 chp2 := hString(Parms[3]);
 valeur2 := hString(Parms[4]);
 libelle := hString(Parms[5]);

 if (valeur1 = '%')
  then
   begin
    HShowMessage('1;Erreur-Contrôle de validité;La valeur de $$ : %% est incorrecte.;W;O;O;O',
                 '', libelle);
    res := false;
   end
  else
   begin
    if valeur1 <> ''
     then
      begin
       code := '';
       if (chp2 <> '')
        then code := 'AND ' + chp2 + ' ="' + valeur2 + '"';

       if not ExisteSQL('SELECT ##TOP 1## '+chp1+'  FROM ' + table + ' WHERE ' +
                        chp1 + ' ="' + valeur1 +
                        '" ' + code)
        then
         begin
          HShowMessage('1;Erreur-Contrôle de validité;La valeur de $$ : %% est incorrecte.;W;O;O;O',
                       valeur1, libelle);
          res := false;
         end;
      end;
   end;
 result := res;
end;

//****************************************************************************

//fonction qui retourne vrai si un champ ne doit pas être vide
function FuncSynChpObligatoire(parms: array of variant; nb: integer): variant;
var res: boolean;
begin
  res := true;
  if (hString(Parms[0]) = '') or (hString(Parms[0]) = '%')
    then
  begin
    HShowMessage('1;Erreur-Contrôle de validité;Vous devez renseigner le champ : $$;W;O;O;O',
      '', hString(parms[1]));
    res := false;
  end;
  result := res;
end;

//****************************************************************************

//retourne vrai si la valeur donnée n'existe pas
function FuncSynBPValeurUnique(parms: array of variant; nb: integer): variant;
var res: boolean;
    valeur, table, chp, libelle : hString;
begin
  res := true;
  valeur := hString(Parms[0]);
  table := hString(Parms[1]);
  chp := hString(Parms[2]);
  libelle := hString(Parms[3]);

  if ExisteSQL('SELECT ##TOP 1## '+chp+' FROM ' + table + ' WHERE ' +
        chp + ' ="' + valeur +'" ')
   then
    begin
     HShowMessage('1;Erreur-Contrôle de validité;La valeur $$ de %% existe déjà.;W;O;O;O',
                  valeur, libelle);
     res := false;
    end;

  result:=res;
end;

//****************************************************************************

//retourne vrai si la valeur donnée n'existe pas
function FuncSynBPSessionUnique(parms: array of variant; nb: integer): variant;
var res: boolean;
    Q:TQuery;
    codeSession,typeNature:hString;
begin
  res := true;
  codeSession := hString(Parms[0]);

  Q:=MOpenSql('SELECT QBS_TYPENATURE FROM QBPSESSIONBP WHERE ' +
              'QBS_CODESESSION="'+codeSession+'"',
              'SynScriptBP (FuncSynBPSessionUnique).',true);
  if not Q.eof
   then
    begin
     if Q.fields[0].asString='1'
      then typeNature:= TraduireMemoire( 'session d''objectif' )
      else typeNature:= TraduireMemoire( 'session de prévision' );
     HShowMessage('1;Erreur-Contrôle de validité;La session $$ (%%) existe déjà.;W;O;O;O',
                  TypeNature,codeSession);
     res := false;
    end;
  ferme(Q);
  result:=res;
end;

//****************************************************************************


//retourne vrai les dates 1 et 2 sont renseignées et si date1 < date2
function FuncSynBPD1infD2(parms: array of variant; nb: integer): variant;
var res: boolean;
    date1,date2 : TDatetime;
    date1S,date2S,libelleD1,libelleD2 : hString;
begin
  res := true;
  date1S:=hString(Parms[0]);
  date2S:=hString(Parms[1]);
  libelleD1:=hString(Parms[2]);
  libelleD2:=hString(Parms[3]);
  date1:= StrToDateTime0(date1S);
  date2:= StrToDateTime0(date2S);
  if (date2<>0) and (date1<>0) and (date2<date1)
   then
    begin
     HShowMessage('1;Erreur-Contrôle de validité;La date $$ doit être inférieure à la date %%.;W;O;O;O',
                  libelleD2, libelleD1);
     res := false;
    end;

  result:=res;
end;

//****************************************************************************

//cherche le max num noeud et l'incremente
function FuncSynBPIncrementeNumNoeud(parms: array of variant; nb: integer): variant;
var session:hString;
begin
 session:=hString(parms[0]);
 result:=BPIncrementenumNoeud(session);
end;

//****************************************************************************

//retourne vrai si les axes donnés sont différents entre eux
function FuncSynBPControleCodeAxeDiff(parms: array of variant; nb: integer): variant;
var codeaxe1,codeaxe2,codeaxe3,codeaxe4,codeaxe5:hString;
    codeaxe6,codeaxe7,codeaxe8,codeaxe9,codeaxe10:hString;
begin
 result:=true;
 codeaxe1:=hString(parms[0]);
 codeaxe2:=hString(parms[1]);
 codeaxe3:=hString(parms[2]);
 codeaxe4:=hString(parms[3]);
 codeaxe5:=hString(parms[4]);
 codeaxe6:=hString(parms[5]);
 codeaxe7:=hString(parms[6]);
 codeaxe8:=hString(parms[7]);
 codeaxe9:=hString(parms[8]);
 codeaxe10:=hString(parms[9]);
 if ((codeaxe1=codeaxe2) and (codeaxe2<>'')) or
    ((codeaxe1=codeaxe3) and (codeaxe3<>'')) or
    ((codeaxe1=codeaxe4) and (codeaxe4<>'')) or
    ((codeaxe1=codeaxe5) and (codeaxe5<>'')) or
    ((codeaxe1=codeaxe6) and (codeaxe6<>'')) or
    ((codeaxe1=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe1=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe1=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe1=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe2=codeaxe3) and (codeaxe3<>'')) or
    ((codeaxe2=codeaxe4) and (codeaxe4<>'')) or
    ((codeaxe2=codeaxe5) and (codeaxe5<>'')) or
    ((codeaxe2=codeaxe6) and (codeaxe6<>'')) or
    ((codeaxe2=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe2=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe2=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe2=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe3=codeaxe4) and (codeaxe4<>'')) or
    ((codeaxe3=codeaxe5) and (codeaxe5<>'')) or
    ((codeaxe3=codeaxe6) and (codeaxe6<>'')) or
    ((codeaxe3=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe3=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe3=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe3=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe4=codeaxe5) and (codeaxe5<>'')) or
    ((codeaxe4=codeaxe6) and (codeaxe6<>'')) or
    ((codeaxe4=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe4=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe4=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe4=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe5=codeaxe6) and (codeaxe6<>'')) or
    ((codeaxe5=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe5=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe5=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe5=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe6=codeaxe7) and (codeaxe7<>'')) or
    ((codeaxe6=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe6=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe6=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe7=codeaxe8) and (codeaxe8<>'')) or
    ((codeaxe7=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe7=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe8=codeaxe9) and (codeaxe9<>'')) or
    ((codeaxe8=codeaxe10) and (codeaxe10<>'')) or
    ((codeaxe9=codeaxe10) and (codeaxe10<>''))
  then
   begin
    result:=false;
    HShowMessage('1;Erreur-Contrôle de validité;Les axes doivent être uniques.;W;O;O;O',
          '', '');
   end;
end;

//****************************************************************************

//retourne vrai la valeur de l'axe donnée existe déjà
function FuncSynBPControleValeurAxeDiff(parms: array of variant; nb: integer): variant;
var valeuraxe,session,numpere:hString;
begin
 result:=true;
 valeuraxe:=hString(parms[0]);
 session:=hString(parms[1]);
 numpere:=hString(parms[2]);
 if ExisteSQL('SELECT QBI_VALEURAXE FROM QBPSAISIEPREV '+
              'WHERE QBI_CODESESSION="'+session+
              '" AND QBI_VALEURAXE="'+valeuraxe+
              '" AND QBI_NUMNOEUDPERE="'+numpere+'"')
  then
   begin
    result:=false;
    HShowMessage('1;Erreur-Contrôle de validité;Données déjà existantes.;W;O;O;O',
          '', '');
   end;
end;

//****************************************************************************

//retourne le libellé du code axe donné
function FuncSynBPDonneLibelleCodeAxe(parms: array of variant; nb: integer): variant;
var codestruct,num:hString;
begin
 codestruct:=hString(parms[0]);
 num:=hString(parms[1]);
 result:=DonneLibelleCodeAxe(codestruct,num);
end;

//****************************************************************************

//retourne le libellé de la valeur axe donnée
function FuncSynBPDonneLibelleValeurAxe(parms: array of variant; nb: integer): variant;
var codeaxe,valeuraxe:hString;
begin
 codeaxe:=hString(parms[0]);
 valeuraxe:=hString(parms[1]);
 result:=DonneLibelleValeurAxe(codeaxe,valeuraxe);
end;

//****************************************************************************

//retourne le code axe suivant le numéro d'axe donné
function FuncSynBPDonneCodeAxe(parms: array of variant; nb: integer): variant;
var codestruct,num:hString;
    Q:TQuery;
begin
 codestruct:=hString(parms[0]);
 num:=hString(parms[1]);
 result:='';
 Q:=MOPenSql('SELECT QBC_CODEAXES'+num+' FROM QBPSTRUCTURE '+
             'WHERE QBC_CODESTRUCT="'+codestruct+'" ',
             'SynScriptBP (FuncSynBPSessionUnique)',true);
 if not Q.eof
  then result:=Q.fields[0].AsString;
 ferme(Q);
end;


//****************************************************************************

function FuncSynBPDonneSessionOk(parms: array of variant; nb: integer): variant;
var Q:TQuery;
    nature,code:hString;
begin
 nature:=hString(parms[0]);
 code:='';
 Q:=OpenSql('SELECT QBS_CODESESSION FROM QBPSESSIONBP,QBPPROFIL '+
            'WHERE QBS_CODEPROFIL=QBF_CODEPROFIL AND QBF_NATURE="'+nature+'"',
            true);
 while not Q.eof do
  begin
   code:=code+' OR QBS_CODESESSION="'+Q.fields[0].asString+'" ';
   Q.next;
  end;
 ferme(Q);
 if code=''
  then code:=' QBS_CODESESSION="" ';
 result:=code;
end;

//****************************************************************************

function FuncSynBPLanceChxValeurAxe(parms: array of variant; nb: integer): variant;
var session,multi,codeaxe,NomTable,ChpCode,ChpLib,Chp2,Val2:hString;
begin
 result:='';
  multi:=hString(parms[0]);
  codeaxe:=hString(parms[1]);
  session:=hString(parms[2]);
  DonneNomTableChampsPrValeurAxe(codeaxe,NomTable,ChpCode,ChpLib,Chp2,Val2);
  Case ContextBP of
    0,1,2,3 : begin
              if codeaxe<>'' then
              {$IFDEF EAGLSERVER}
              result := ''
              {$ELSE  EAGLSERVER}
              result:=AglLanceFiche('Q', 'QUFMBPMUL', '', '','CODEAXE='+codeaxe+';MULTI='+multi+';SESSION='+session)
              {$ENDIF EAGLSERVER}
            end;
  else //CASE
    {$IFNDEF EAGLCLIENT}
    {$IFDEF BPORLI}
     //-----------------> ORLI
     if NomTable<>'' then result:=OuvertureFicheValeurAxe(multi,codeaxe,NomTable,ChpCode,ChpLib,Chp2,Val2);
    //ORLI <-----------------
    {$ENDIF}
    {$ENDIF}
  end; //CASE
end;


//****************************************************************************

function FuncSynBPInitialiseDelaiOk(parms: array of variant; nb: integer): variant;
var codesession,_:hString;
begin
 codesession:=hString(parms[0]);
 result:=SessionCalculParDelai(codesession,_)
end;


//****************************************************************************

//initialise la table des valeurs axes
procedure ProcSynBPInitialise(parms: array of variant; nb: integer);
var codeSession:hString;
  //  rep:integer;
begin
 codeSession:=hString(parms[0]);
 //rep:=HShowmessage('1;Initialisation du budget.;Etes vous sûr ?;Q;YN;N;N', '', '');
// if rep = mrYes
//  then InitialiseTableQbpSaisiePrev(codeSession);
end;


//****************************************************************************

function FuncSynBPPGIOrli(parms: array of variant; nb: integer): variant;
begin
 result:=not BPOkOrli;
end;

function FuncSynBPOkOrli(parms: array of variant; nb: integer): variant;
begin
 result:=BPOkOrli;
end;

function FuncSynBPOkPgiEnt(parms: array of variant; nb: integer): variant;
begin
 if contextBP in [0,2,3] then result:=true else result:=false;
end;


//****************************************************************************

function FuncSynDonneLibCodeAxeBP(parms: array of variant; nb: integer): variant;
var codeAxe:hString;
begin
 result:='';
 codeAxe:=(parms[0]);
 result:=DonneLibelleAxe(codeAxe);
end;

//****************************************************************************

//retourne le code pour la tablette
//pour avoir les différents codes axes de pivot
//qui existent dans la structure donnée
function FuncSynBPDonneCodeTablette(parms: array of variant; nb: integer): variant;
var codestruct:hString;
begin
 codestruct:=hString(parms[0]);
 result:='';
 if BPOkOrli
  then //-----------------> ORLI
       result:=DonneCodeTabletteAxeStructure(codestruct);
       //ORLI <-----------------
end;

//****************************************************************************

function FuncSynBPValeurAxePlus(parms: array of variant; nb: integer): variant;
var codeaxe:hString;
begin
 codeaxe:=hString(parms[0]);
 if BPOkOrli
  //-----------------> ORLI
  then result:='QBA_CODEAXE="'+codeaxe+'"'
  //ORLI <-----------------
  else result:='';
end;

//****************************************************************************

function FuncSynBPControleValiditeCodeAxe(parms: array of variant; nb: integer): variant;
var codeaxe,structure,msg:hString;
begin
 result:=true;
 codeaxe:=hString(parms[0]);
 structure:=hString(parms[1]);
 msg:=hString(parms[2]);
 result:=ControleAxe(codeAxe,structure,msg);
end;

//****************************************************************************

function FuncSynBPControleValiditeValeurAxe(parms: array of variant; nb: integer): variant;
var codeaxe,valeuraxe,msg:hString;
begin
 result:=true;
 valeuraxe:=hString(parms[0]);
 codeaxe:=hString(parms[1]);
 msg:=hString(parms[2]);
 result:=ControleValeurAxe(codeAxe,valeurAxe,msg);
end;

//****************************************************************************

function FuncSynBPTabletteCodeAxe(parms: array of variant; nb: integer): variant;
begin
 {$IFNDEF CRM}
 Case contextBP of
   0 : result:='QUTBPAXEENT';
   1 : result:='QUTBPAXE';
   2 : result:='QUTBPAXECOMPTA';
   3 : result:='QUTBPAXEPAIE';
 end;
 {$ELSE CRM}
 if (GetParamSocSecur ('SO_CRMACCOMPAGNEMENT', False) = True) then result:='QUTBPAXEENT'
 else result:='QUTBPAXECRM';
 {$ENDIF}
end;

//****************************************************************************

function FuncSynBPTablettePlus(parms: array of variant; nb: integer): variant;
var codeStruct:hString;
begin
 codeStruct:=hString(parms[0]);
 result:=BPTablettePlus(codeStruct);
end;

//****************************************************************************

function FuncSynBPTablettePlusSession(parms: array of variant; nb: integer): variant;
var TabCodeAxe:array [1..11] of hString;
    i:integer;
    code:hString;
begin
 result:='';
 if BPOkOrli
  then
   //-----------------> ORLI
   begin
    ChercheCodeAxeSession(hString(parms[0]),TabCodeAxe);
    code:='';
    for i:=1 to 10 do
     code:=code+' OR QBX_CODEAXE="'+TabCodeAxe[i]+'" ';
    result:=code;
   end
   //ORLI <-----------------
  else result:='QBX_CODEMARCHE="#PGI"';
end;


function FuncSynBPSessionCodeAxe(parms: array of variant; nb: integer): variant;
var TabCodeAxe:array [0..11] of hString;
    i:integer;
begin
 result:='';
 ChercheCodeAxeSession(hString(parms[0]),TabCodeAxe);
 for i:=1 to 10 do
  result:=result+IntTostr(i)+':='+TabCodeAxe[i]+';';
end;

//****************************************************************************

function FuncSynBPSessionNivMax(parms: array of variant; nb: integer): variant;
begin
 result:=ChercheNivMax(hString(parms[0]));
end;
//****************************************************************************

function FuncSynBPPlusValeurAxe(parms: array of variant; nb: integer): variant;
var codeAxe:hString;
begin
 result:='';
 codeAxe:=hString(parms[0]);
 if BPOkOrli
  then //-----------------> ORLI
       result:='QBA_CODEAXE="'+codeAxe+'"';
       //ORLI <-----------------
end;

//****************************************************************************

function FuncSynBPTabletteValeurAxe(parms: array of variant; nb: integer): variant;
begin
 result:='';
 if BPOkOrli
  then //-----------------> ORLI
       result:='QUTBPVALEURAXE'
       //ORLI <-----------------
  else result:='';
end;

function FuncSynBPTabletteValeurAxeC(parms: array of variant; nb: integer): variant;
var codeAxe:hString;
begin
 result:='';
 codeAxe:=hString(parms[0]);
 if BPOkOrli
  then //-----------------> ORLI
       result:='QUTBPVALEURAXE'
       //ORLI <-----------------
  else result:=ChercheTabletteValeurAxe(codeAxe);
end;



//****************************************************************************

function FuncSynBPTabletteValeurAxeSaiCmd(parms: array of variant; nb: integer): variant;
begin
 result:='';
 if BPOkOrli
  then //-----------------> ORLI
       result:='QUTBPVALEURAXE'
       //ORLI <-----------------
  else result:='';
end;


//****************************************************************************

function FuncSynBPTabPlusValeurAxeSaiCmd(parms: array of variant; nb: integer): variant;
var codeaxe:hString;
begin
 result:='';
 if BPOkOrli
  then
   //-----------------> ORLI
   begin
    codeAxe:='SAISONCMD';
    result:=' QBA_CODEAXE="'+codeAxe+'" ';
   end
   //ORLI <-----------------
  else result:='';
end;

//****************************************************************************

//fonction qui retourne le code session incrementé
//pour la duplication d'une session
//si nom sans _00X derriere => ajoute au nom _001
//sinon incremente le num 00X
function FuncSynBPIncrementeSession(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 codeSession:=hString(parms[0]);
 result:=IncrementeCodeSession(codeSession);
end;

//****************************************************************************

//procedure qui lance le calcul des coefficients d'extrapolation
//remplit la table qbpcoeff à partir de l'historique
//pour la session donnée
procedure ProcSynBPLanceCalculInitCoeff(parms: array of variant; nb: integer);
var codeSession,BPInitialise:hString;
    rep:integer;
begin
 codeSession:=hString(parms[0]);
 BPInitialise:=hString(parms[1]);
 if SessionValide(codeSession)
  then
   begin
    PGIINFO('Une session validée ne peut être modifiée.');
    exit;
   end;
 rep:=HShowmessage('1;Initialisation des coefficients.;' + Format_( TraduireMemoire( TexteMessage[ 5 ] ), [ codeSession, #13#10 ] ) + ';Q;YN;N;N','','');
 if rep = mrYes
  then IntialiseCoeff(codeSession,'0',false);
end;

//****************************************************************************

//procedure qui lance le calcul du budget
//remplit la table qbpsaisieprev (càd l'arbre)
//pour la session donnée
//par délai à partir de l'historique
//càd rajoute dans l'arbre un niveau supplémentaire pour les délais
Function FuncSynBPLanceCalculInitObjectif(parms: array of variant; nb: integer):variant;
var codeSession,codeMaille,SaiRef,SaiC:hString;
    rep,i:integer;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    NbMailleC,NbMailleRef:double;
    aD,m,d,aF:word;
    OkInit:boolean;
begin
  result := false;
  codeSession:=hString(parms[0]);
  codeMaille:=hString(parms[1]);
  SaiC:=hString(parms[2]);
  SaiRef:=hString(parms[3]);
  DateDebC:=StrToDateTime(hString(parms[4]));
  DateFinC:=StrToDateTime(hString(parms[5]));
  DateDebRef:=StrToDateTime(hString(parms[6]));
  DateFinRef:=StrToDateTime(hString(parms[7]));
  if SessionValide(codeSession) then
  begin
    PGIINFO('Une session validée ne peut être modifiée.');
    exit;
  end;

  if (codeMaille='0') or (codeMaille='') then
  begin
    //test nombre de mailles égales
    NbMailleC:=DateFinC-DateDebC;
    Decodedate(DateDebC,aD,m,d);
    Decodedate(DateFinC,aF,m,d);

    for i:=aD to aF do
    begin
      if IsLeapYear(i) then
      begin
        { EVI / Appel année (aD) -> année courante (i) IsLeapYear }        if DateAppartIntervalle(encodedate(i,2,29),DateDebC,DateFinC) then NbMailleC:=NbMailleC-1;
      end;
    end;

    NbMailleRef:=DateFinRef-DateDebRef;
    Decodedate(DateDebRef,aD,m,d);
    Decodedate(DateFinRef,aF,m,d);

    for i:=aD to aF do
    begin
      if IsLeapYear(i) then
      begin
        { EVI / Appel année (aD) -> année courante (i) IsLeapYear }
        if DateAppartIntervalle(encodedate(i,2,29),DateDebRef,DateFinRef) then NbMailleRef:=NbMailleRef-1;
      end;
    end;

    OkInit:=true;
    if (SaiC='') and (SaiRef='') then
    if (NbMailleC<>NbMailleRef) then
    begin
      PGIINFO(Format_( TraduireMemoire( TexteMessage[ 6 ] ), [ #13#10, FloatToStr(NbMailleC), FloatTostr(NbMailleRef) ] ) );
      OkInit:=false;
    end;

    { EVI / Contrôle session Loi d'Eclatement inférieure à 2 ans }
    if (SaiC='') and (SaiRef='') then
    if (NbMailleC > 732) OR (NbMailleRef > 732) then
    begin
      PGIINFO(Format_( TraduireMemoire( TexteMessage[ 7 ] ), [ #13#10 ] ) +
              Format_( TraduireMemoire( TexteMessage[ 8 ] ), [ FloatToStr(NbMailleC), FloatTostr(NbMailleRef)] ) );
      OkInit:=false;
    end;

    if OkInit then
    begin
      //intialise en global
      rep:=HShowmessage('1;Vous allez initialiser la session en global.;' + Format_( TraduireMemoire(  TexteMessage[ 5 ] ), [ codeSession, #13#10 ] )+  ';Q;YN;N;N', '', '');
      if rep = mrYes then
      begin
        if QuestBPInit(codeSession) then
        if QuestBPCalend(codeSession) then
        begin
          InitialiseGlobal(codeSession);
          result:=true;
        end;
      end;
    end;
  end
  else result := InitialiseDelai(codeSession,codeMaille,SaiRef,SaiC,DateDebC,DateFinC,DateDebRef,DateFinRef);
end;


//****************************************************************************

procedure ProcSynBPLanceCalculEclatementLoi(parms: array of variant; nb: integer);
var codeSession,TypeEclat,codeSql,codeMsg:hString;
    i,rep:integer;
    okCalcul:boolean;
    DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime;
    NbMailleC,NbMailleRef:double;
    aD,m,d,aF:word;
begin
  codeSession:=hString(parms[0]);
  TypeEclat:=hString(parms[1]);
  DateDebC:=StrToDateTime(hString(parms[2]));
  DateFinC:=StrToDateTime(hString(parms[3]));
  DateDebRef:=StrToDateTime(hString(parms[4]));
  DateFinRef:=StrToDateTime(hString(parms[5]));

  if SessionValide(codeSession) then
  begin
    PGIINFO('Une session validée ne peut être modifiée.');
    exit;
  end;

  if (not OkSessionLoiDefinie(codeSession)) and (ContextBP=1) then
  begin
    PGIINFO('Il n''y a pas de lois d''éclatement définies pour cette session.');
    exit;
  end;

  //test nombre de mailles égales
  NbMailleC:=DateFinC-DateDebC;
  Decodedate(DateDebC,aD,m,d);
  Decodedate(DateFinC,aF,m,d);

  for i:=aD to aF do
  begin
    if IsLeapYear(i) then
    begin
      if DateAppartIntervalle(encodedate(i,2,29),DateDebC,DateFinC)
      then NbMailleC:=NbMailleC-1;
    end;
  end;

  NbMailleRef:=DateFinRef-DateDebRef;
  Decodedate(DateDebRef,aD,m,d);
  Decodedate(DateFinRef,aF,m,d);

  for i:=aD to aF do
  begin
    if IsLeapYear(i) then
    begin
      if DateAppartIntervalle(encodedate(i,2,29),DateDebRef,DateFinRef)
      then NbMailleRef:=NbMailleRef-1;
    end;
  end;

  if (NbMailleC<>NbMailleRef) then
  begin
    PGIINFO(Format_( TraduireMemoire( TexteMessage[ 6 ] ), [ #13#10, FloatToStr(NbMailleC), FloatTostr(NbMailleRef) ] ) );
    exit;
  end;

  if (NbMailleC > 732) OR (NbMailleRef > 732) then
  begin
    PGIINFO(Format_( TraduireMemoire( TexteMessage[ 7 ] ), [ #13#10 ] ) +
            Format_( TraduireMemoire( TexteMessage[ 8 ] ), [ FloatToStr(NbMailleC), FloatTostr(NbMailleRef)] ) );
    exit;
  end;

  if ContextBP <> 1
  then
   begin
     { EVI / Rectification du code SQL qui ne vérifiait pas la présence de loi
     pour la session donnée mais pour nimporte quelle session :
     codeSql:='SELECT QBO_CODELOI FROM QBPLOI'; }
     codeSql:='SELECT QBO_CODELOI FROM QBPLOI WHERE QBO_CODESESSION="'+codeSession+'"';
     codeMsg:=traduireMemoire('Il faut definir des lois d''éclatement par délai.');

     if TypeEclat='TAILLE'
      then
       begin
        codeSql:='SELECT QBL_CODELOITAILLE FROM QBPLOITAILLE';
        codeMsg:=traduireMemoire('Il faut definir des lois d''éclatement par taille.');
       end;

    if not ExisteSql(codeSql)
     then
      begin
       PGIINFO(codeMsg);
       exit;
      end;
   end;

 if (TypeEclat='TAILLE') and (not SessionDelai(codeSession))
  then
   begin
    PGIINFO('Le calcul d''éclatement par taille ne peut se faire qu''à partir d''une session definie par délai.');
    exit;
   end;

 okCalcul:=false;
 if TypeEclat='TAILLE'
  then
   begin
    rep:=HShowmessage('1;Eclatement par taille.;Etes-vous sûr ?;Q;YN;N;N', '', '');
     if rep=mrYes
     then
      begin
       if SessionEclateeParTaille(codeSession)
        then OkCalcul:=QuestionBP()
        else OkCalcul:=true;
      end;
   end
  else
   begin
    rep:=HShowmessage('1;Eclatement par délai.;Etes-vous sûr ?;Q;YN;N;N', '', '');
    if rep=mrYes
     then
      begin
       if SessionEclateeParDelai(codeSession)
        then OkCalcul:=QuestionBP()
        else OkCalcul:=true;
      end;
   end;

 if OkCalcul
  then CalculEclatementLoi(codeSession,TypeEclat);

end;

//****************************************************************************

procedure ProcSynBPLanceValidation(parms: array of variant; nb: integer);
var codeSession:hString;
    rep:integer;
begin
 codeSession:=hString(parms[0]);
 if SessionValide(codeSession)
  then
   begin
    PGIINFO('Une session validée ne peut être modifiée.');
    exit;
   end;
 rep:=HShowmessage('1;Une fois validée, la session ne sera plus modifiable.;Confirmez-vous la validation ?;Q;YN;N;N', '', '');
 if rep = mrYes
  then ValidationBP(codeSession);
end;

//****************************************************************************

procedure ProcSynBPLanceDeValidation(parms: array of variant; nb: integer);
var codeSession:hString;
    rep:integer;
begin
 codeSession:=hString(parms[0]);
 rep:=HShowmessage('1;Une fois dévalidée, la session sera à nouveau modifiable.;Confirmez-vous la dévalidation ?;Q;YN;N;N', '', '');
 if rep = mrYes
  then  MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONVALIDE="-",'+
             'QBS_DATEVALIDATION="'+USDATETIME(idate1900)+
             '" WHERE QBS_CODESESSION="'+codeSession+'"',
             'ProcSynBPLanceDeValidation (SynScriptBP).');
end;

//****************************************************************************

procedure ProcSynBPValidationVersSCM(parms: array of variant; nb: integer);
var codeSession:hString;
    rep:integer;
begin
 codeSession:=hString(parms[0]);
 rep:=HShowmessage('1;Alimenter la SCM.;Etes-vous sûr ?;Q;YN;N;N', '', '');
 if rep = mrYes
  then CalculPrevCmd(codeSession,'');//ValidationBPdansSCM(codeSession);
end;

//****************************************************************************

function FuncSynBPDonneLibelleSession(parms: array of variant; nb: integer): variant;
var codeSession:hString;
    Q:TQuery;
begin
 result:='';
 codeSession:=hString(parms[0]);
 Q:=MOpenSql('SELECT QBS_LIBSESSION FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'"',
             'FuncSynBPDonneLibelleSession (SynScriptBP).',true);
 if not Q.eof
  then result:=Q.fields[0].asString;
 ferme(Q);
end;

//****************************************************************************

function FuncSynBPArbreControleValeurAxe(parms: array of variant; nb: integer): variant;
var codeSession,noeudPere,codeaxe,valeuraxe:hString;
    NomTable,chpCode,ChpLib,Chp2,Val2,msg:hString;
begin
  result:=true;
  codeSession:=hString(parms[0]);
  noeudPere:=hString(parms[1]);
  codeaxe:=hString(parms[2]);
  valeuraxe:=hString(parms[3]);
  msg:='';
  if valeuraxe<>'' then
  begin
    //contrôle si valeuraxe est correct
    DonneNomTableChampsPrValeurAxe(codeaxe,NomTable,chpCode,ChpLib,Chp2,Val2);
    if not ExisteSql('SELECT '+chpCode+' FROM '+NomTable+
                     ' WHERE '+chpCode+'="'+valeurAxe+'"')
    then
    begin
      msg:= Format_( TraduireMemoire( 'La valeur axe entrée est incorrecte pour le code axe %s. ' ), [codeaxe] );
      result:=false;
    end;

    //contrôle si valeuraxe n'existe pas déjà dans l'arbre
    if ExisteSql('SELECT QBR_CODESESSION FROM QBPARBRE '+
                    'WHERE QBR_CODESESSION="'+codeSession+
                    '" AND QBR_CODEAXE="'+codeAxe+
                    '" AND QBR_VALEURAXE="'+valeurAxe+
                    '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"')
    then
    begin
      msg:=Format_( TraduireMemoire( 'La valeur axe entrée existe déjà pour le code axe %s. ' ), [codeaxe] );
      result:=false;
    end;
  end
  else
  begin
    msg:= TraduireMemoire( 'La valeur axe n''est pas renseignée.' );
    result:=false;
  end;
  if result=false then PGIINFO(msg);
end;

//****************************************************************************

function FuncSynBPArbreControleDate(parms: array of variant; nb: integer): variant;
var codeSession,noeudPere,codeaxe,valeuraxe,msg:hString;
begin
  result:=true;
  codeSession:=hString(parms[0]);
  noeudPere:=hString(parms[1]);
  codeaxe:=hString(parms[2]);
  valeuraxe:=hString(parms[3]);
  msg:='';
  if valeuraxe<>'' then
  begin
    //contrôle si valeuraxe n'existe pas déjà dans l'arbre
    if ExisteSql('SELECT QBR_CODESESSION FROM QBPARBRE '+
                    'WHERE QBR_CODESESSION="'+codeSession+
                    '" AND QBR_VALEURAXE="'+valeurAxe+
                    '" AND QBR_NUMNOEUDPERE="'+noeudPere+'"')
    then
    begin
      msg := Format_( TraduireMemoire( 'La date %s existe déjà. '), [valeuraxe] );
      result:=false;
    end;
  end
  else
  begin
    msg := TraduireMemoire( 'La valeur axe n''est pas renseignée.' );
    result:=false;
  end;
  if result=false then PGIINFO(msg);
end;

//****************************************************************************

function FuncSynBPDonneDateDebSaison(parms: array of variant; nb: integer): variant;
var Saison:hString;
    DateDeb,DateFin:TDateTime;
begin
 result:='';
 Saison:=hString(parms[0]);
 DonneDateDebEtFinSaison(Saison,DateDeb,DateFin);
 result:=DateTimeToStr(DateDeb);
end;

function FuncSynBPDonneDateFinSaison(parms: array of variant; nb: integer): variant;
var Saison:hString;
    DateDeb,DateFin:TDateTime;
begin
 result:='';
 Saison:=hString(parms[0]);
 DonneDateDebEtFinSaison(Saison,DateDeb,DateFin);
 result:=DateTimeToStr(DateFin);
end;

//****************************************************************************




//****************************************************************************

function FuncSynBPDonneDateDebSession(parms: array of variant; nb: integer): variant;
var codeSession:hString;
    DDC,DFC,DDR,DFR:TDatetime;
begin
 result:='';
 codeSession:=hString(parms[0]);
 ChercheDateDDateFPeriode(codeSession,DDC,DFC,DDR,DFR);
 result:=DateTimeToStr(DDC);
end;

function FuncSynBPDonneDateFinSession(parms: array of variant; nb: integer): variant;
var codeSession:hString;
    DDC,DFC,DDR,DFR:TDatetime;
begin
 result:='';
 codeSession:=hString(parms[0]);
 ChercheDateDDateFPeriode(codeSession,DDC,DFC,DDR,DFR);
 result:=DateTimeToStr(DFC);
end;

//****************************************************************************

procedure ProcSynBPMAJAxeSaison(parms: array of variant; nb: integer);
var saison,saisonLib:hString;
begin
 saison:=hString(parms[0]);
 saisonLib:=hString(parms[1]);
 if not ExisteSql('SELECT QBA_CODEAXE FROM QBPAXE WHERE '+
                 'QBA_CODEAXE="'+'SAISONCMD'+'" AND QBA_VALEURAXE="'+saison+'"')
  then MExecuteSql('INSERT INTO QBPAXE (QBA_CODEAXE,QBA_VALEURAXE,'+
                   'QBA_LIBVALEURAXE) VALUES ("'+'SAISONCMD'+
                   '","'+saison+'","'+saisonLib+'")',
                   'SynScriptBP (ProcSynBPMAJAxeSaison).')
  else MExecuteSql('UPDATE QBPAXE SET QBA_LIBVALEURAXE="'+saisonLib+
                   '" WHERE QBA_CODEAXE="'+'SAISONCMD'+
                   '" AND QBA_VALEURAXE="'+saison+'")',
                   'SynScriptBP (ProcSynBPMAJAxeSaison).');
end;

//****************************************************************************

function FuncSynBPDonneMethodeSession(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:='';
 codeSession:=hString(parms[0]);
 result:=DonneMethodeSession(codeSession);
end;

//****************************************************************************

procedure ProcSynBPExportVersOrliWeb(parms: array of variant; nb: integer);
var codeSession:hString;
begin
 codeSession:=hString(parms[0]);
 //LanceExportationPrevVersOrliWeb(codeSession);
end;

procedure ProcSynBPExportVersOrliPap(parms: array of variant; nb: integer);
var codeSession:hString;
begin                                      
 codeSession:=hString(parms[0]);
 //LanceExportationPrevVersOrliPap(codeSession);
end;

//****************************************************************************

function FuncSynBPDateDebutAnnee(parms: array of variant; nb: integer): variant;
begin
 result:=DEBUTANNEE(date);
end;

function FuncSynBPDateFinAnnee(parms: array of variant; nb: integer): variant;
begin
 result:=FINANNEE(date);
end;

function FuncSynBPDateJour(parms: array of variant; nb: integer): variant;
begin
 result:=date;
end;

function FuncSynBPDateDebutAnneeMoinsUn(parms: array of variant; nb: integer): variant;
begin
 result:=DEBUTANNEE(date-365);
end;

function FuncSynBPDateFinAnneeMoinsUn(parms: array of variant; nb: integer): variant;
begin
 result:=FINANNEE(date-365);
end;

//****************************************************************************


function FuncSynBPSessionBudgetOk(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 codeSession:=hString(parms[0]);
 result:=OkSessionObjectif(codeSession);
end;

//****************************************************************************


function FuncSynBPSessionPermanent(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 codeSession:=hString(parms[0]);
 result:=OkSessionPermanent(codeSession);
end;

//****************************************************************************

function FuncSynBPDonneTabVisible(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 codeSession:=hString(parms[0]);
 result:=DonneValeurAffiche(CodeSession);
end;

//****************************************************************************

procedure ProcSynBPDuplicationArbre(parms: array of variant; nb: integer);
var codeSession,codeSessionN:hString;
begin
 codeSession:=hString(parms[0]);
 codeSessionN:=hString(parms[1]);
 DuplicationArbreBP(codeSession,codeSessionN);
end;

procedure ProcSynBPDuplicationArbrePartielle(parms: array of variant; nb: integer);
var codeSession,codeSessionN,Niv:hString;
begin
 codeSession:=hString(parms[0]);
 codeSessionN:=hString(parms[1]);   
 NIv:=hString(parms[2]);
 DuplicationArbrePartielleBP(codeSession,codeSessionN,Niv);
end;

procedure ProcSynBPDuplicationArbreBP(parms: array of variant; nb: integer);
var codeSession,codeSessionN:hString;
begin
 codeSession:=hString(parms[0]);
 codeSessionN:=hString(parms[1]);
 DuplicationArbreBP(codeSession,codeSessionN);
end;

procedure ProcSynBPDuplicationArbreBPPartielle(parms: array of variant; nb: integer);
var codeSession,codeSessionN,Niv:hString;
begin
 codeSession:=hString(parms[0]);
 codeSessionN:=hString(parms[1]);
 NIv:=hString(parms[2]);
 DuplicationArbrePartielleBP(codeSession,codeSessionN,Niv);
end;

procedure ProcSynBPLanceCalculInitPrevision(parms: array of variant; nb: integer);
var codeSession:hString;
    rep:integer;
begin
 codeSession:=hString(parms[0]);
 if SessionValide(codeSession)
  then
   begin
    PGIINFO('Une session validée ne peut être modifiée.');
    exit;
   end;
 if OkSessionObjectif(codeSession)
  then
   begin
    PGIINFO('Ne peut initialiser une session d''objectif.');
    exit;
   end;
 rep:=HShowmessage('1;Initialisation de la prévision.;' + Format_( TraduireMemoire( TexteMessage[ 5 ] ), [ codeSession, #13#10 ] ) + ';Q;YN;N;N', '', '');
 if rep = mrYes
  then
   begin
    if not SessionInitCoeff(codeSession)
     then IntialiseCoeff(codeSession,'0',true);
    BPIntialisePrevision(codeSession);
   end;
end;

//****************************************************************************

procedure ProcSynBPModifSessionVue(parms: array of variant; nb: integer);
var codeSession,codeM,ModifChpSession:hString;
begin
 codeSession:=hString(parms[0]);
 codeM:=hString(parms[1]);       
 ModifChpSession:=hString(parms[2]);
 ModifSessionVue(codeSession,codeM,ModifChpSession);
end;



//****************************************************************************


function FuncSynBPBtnInitGlobalObj(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if (OkSessionObjectif(codeSession)=true) and
    (DonneMethodeSession(codeSession)='2')
  then result:=true;
end;

function FuncSynBPBtnInitDelaiObj(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if (OkSessionObjectif(codeSession)=true) and
    (DonneMethodeSession(codeSession)='1')
  then result:=true;
end;

function FuncSynBPBtnInitCoeff(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if (OkSessionObjectif(codeSession)=false)
  then result:=true;
end;

function FuncSynBPBtnInitPrev(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 result:=SessionInitPrev(codeSession);
end;

function FuncSynBPBtnInitPrevDelai(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if (OkSessionObjectif(codeSession)=false) and
    (DonneMethodeSession(codeSession)='1')
  then result:=true;
end;


function FuncSynBPBtnInitPrevTaille(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if (OkSessionObjectif(codeSession)=false) and
    (SessionCalculParTaille(codeSession))
  then result:=true;
end;

function FuncSynBPSessionValide(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 if SessionValide(codeSession) or ObjPrevConsultant()
  then result:=true;
end;

// GC_20080124_GM_GC15727 debut
function FuncSynBPSessionDelete(parms: array of variant; nb: integer): variant;
var
{$IFNDEF MODE}
  {$IFDEF GCGC}
  scodeSession:hString;
  sSessionN,sSessionN1 :String;
  {$ENDIF}
{$ENDIF}
begin
 result:=True;
{$IFNDEF MODE}
  {$IFDEF GCGC}
    scodeSession:=hString(parms[0]);
    sSessionN := GetParamsocSecur('SO_GCCODESESSIONN','');
    sSessionN1 := GetParamsocSecur('SO_GCCODESESSIONN1','');
    // GC_20080430_JPL_GC15973_DEBUT
//    if (sSessionN = scodeSession) or ( sSessionN1 = scodeSession)
//    or (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + sCodeSession +'"')) then
    if (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + sCodeSession +'"')) then
    // GC_20080430_JPL_GC15973_FIN
    Begin
      PGIINFO('Vous ne pouvez pas supprimer une session attachée aux budgets pour les DA');
      Result := False;
    End;
  {$ENDIF}
{$ENDIF}
end;
// GC_20080124_GM_GC15727 fin

procedure ProcSynBPSessionMAJ(parms: array of variant; nb: integer);
var codeSession:hString;
begin
 if ContextBP = 1 then
 begin
   codeSession:=hString(parms[0]);
   SessionMAJ(codeSession);
 end;
end;

function FuncSynBPOkDelai(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 result:=SessionDelai(codeSession);
end;


function FuncSynBPOkTaille(parms: array of variant; nb: integer): variant;
var codeSession:hString;
begin
 result:=false;
 codeSession:=hString(parms[0]);
 result:=SessionCalculParTaille(codeSession);
end;



//****************************************************************************

function FuncSynBPIncrementeNumSession(parms: array of variant; nb: integer): variant;
var Q:TQuery;
begin
 Q:=MOpenSql('SELECT MAX(QBS_NUMSESSION) FROM QBPSESSIONBP ',
             'SynScriptBP (FuncSynBPIncrementeNumSession).',true);
 if not Q.eof
  then result:=Q.fields[0].asInteger+1;
 ferme(Q);
end;

//****************************************************************************

function FuncSynBPValRestrictionValaxe1(parms: array of variant; nb: integer): variant;
var valRestriction,valaxe1,valaxe2,valaxe3,valaxe4:hString;
begin
 valRestriction:=hString(parms[0]);
 ValRestrictionValaxe1234(valRestriction,valaxe1,valaxe2,valaxe3,valaxe4);
 result:=valaxe1;
end;

function FuncSynBPValRestrictionValaxe2(parms: array of variant; nb: integer): variant;
var valRestriction,valaxe1,valaxe2,valaxe3,valaxe4:hString;
begin
 valRestriction:=hString(parms[0]);
 ValRestrictionValaxe1234(valRestriction,valaxe1,valaxe2,valaxe3,valaxe4);
 result:=valaxe2;
end;

function FuncSynBPValRestrictionValaxe3(parms: array of variant; nb: integer): variant;
var valRestriction,valaxe1,valaxe2,valaxe3,valaxe4:hString;
begin
 valRestriction:=hString(parms[0]);
 ValRestrictionValaxe1234(valRestriction,valaxe1,valaxe2,valaxe3,valaxe4);
 result:=valaxe3;
end;

function FuncSynBPValRestrictionValaxe4(parms: array of variant; nb: integer): variant;
var valRestriction,valaxe1,valaxe2,valaxe3,valaxe4:hString;
begin
 valRestriction:=hString(parms[0]);
 ValRestrictionValaxe1234(valRestriction,valaxe1,valaxe2,valaxe3,valaxe4);
 result:=valaxe4;
end;

function FuncSynBPValaxe1234ValRestriction(parms: array of variant; nb: integer): variant;
var valaxe1,valaxe2,valaxe3,valaxe4:hString;
begin
 valaxe1:=hString(parms[0]);
 valaxe2:=hString(parms[1]);
 valaxe3:=hString(parms[2]);
 valaxe4:=hString(parms[3]);
 result:=Valaxe1234ValRestriction(valaxe1,valaxe2,valaxe3,valaxe4);
end;

function FuncSynBPOkListeSem(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeSem);
end;

function FuncSynBPOkListeQuinz(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeQuniz);
end;

function FuncSynBPOkListeMois(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeMois);
end;

function FuncSynBPOkListeMois445(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeMois445);
end;

function FuncSynBPOkListeTrim(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeTrim);
end;

function FuncSynBPOkListeQuadri(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPListeQuadri);
end;

procedure ProcSynBPMiseAJourTabletteMaille(parms: array of variant; nb: integer);
var maille,coche:hString;
begin
 maille:=hString(parms[0]);
 coche:=hString(parms[1]);
 ExecuteSql('UPDATE COMMUN SET CO_ABREGE="'+coche+
            '" WHERE CO_TYPE="QBM" AND CO_CODE="'+maille+'" ');
 //enleve le choix jour pour l'instant
 ExecuteSql('UPDATE COMMUN SET CO_ABREGE="-" WHERE CO_TYPE="QBM" AND CO_CODE="1" ');         
end;
//****************************************************************************

procedure ProcSynBPMAJDatesLoi(parms: array of variant; nb: integer);
var codeSession:hString;
    DateDeb,DateFin:TDateTime;
begin
 codeSession:=hString(parms[0]);
 DateDeb:=StrToDateTime(hString(parms[1]));
 DateFin:=StrToDateTime(hString(parms[2]));

 ExecuteSql('UPDATE QBPLOI SET QBO_DATEDEBECLAT="'+USDATETIME(DateDeb)+
            '",QBO_DATEFINECLAT="'+USDATETIME(DateFin)+
            '" WHERE QBO_CODESESSION="'+codeSession+'" ');

end;

//****************************************************************************

function FuncSynBPLibelleMaille(parms: array of variant; nb: integer): variant;
var BPinitialise,code:hString;
begin
 code:='';
 BPinitialise:=hString(parms[0]);
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

function FuncSynBBOkLoi(parms: array of variant; nb: integer): variant;
var CodeSession,methode:hString;
    Q:TQuery;
begin
 CodeSession:=hString(parms[0]);
 Q:=MOPenSql('SELECT QBS_METHODE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codeSession+'" ',
             'SynScriptBP (FuncSynBBOkLoi).',true);
 if not Q.eof
  then methode:=Q.fields[0].asString;
 ferme(Q);
 result:=(methode='2');
end;        

//****************************************************************************

function FuncSynBBHShowmessage(parms: array of variant; nb: integer): variant;
var CodeSession:hString;
begin
 CodeSession:=hString(parms[0]);
 result := HShowmessage('1;' + Format_( TraduireMemoire( TexteMessage[ 9 ] ), [ CodeSession ] ) + ';' + TexteMessage[ 10 ] + ';Q;YN;N;N', '', '');
end;

//****************************************************************************

function FuncSynBPExportSessionValide(parms: array of variant; nb: integer): variant;
begin
 result:=DonneParamB(pb_BPExportSessionValide);
end;

//****************************************************************************

function FuncSynGestTaille(parms: array of variant; nb: integer): variant;
begin
  result := DonneParamB(pb_GestTaille);
end;

{ EVI / Calcul du nombre de maille pour InitMoveProgressForm }
function NbDeMaille(session,BPinit:hString;DateDeb,DateFin:TDateTime):integer;
var methode:hString;
Q : TQuery;
begin
Q :=OpenSQL('SELECT QBS_DATEDEBREF,QBS_DATEFINREF,QBS_BPINITIALISE FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',True)  ;
  If Not Q.Eof Then
  begin
    if DateDeb=0 then DateDeb:=Q.Fields[0].AsDateTime ;
    if DateFin=0 then DateFin:=Q.Fields[1].AsDateTime ;
    if BPinit='InitOK' then methode:=Q.Fields[2].AsString else methode:=BPinit;
  end;
  Ferme(Q) ;

  { EVI / Contrôle pour session non initialisée }
  if methode='' then methode:='0';

  if BPinit='InitOK' then
  begin
   if methode = '0' then methode := SessionMinLoi(session);
  end;

  { EVI / Session initialisée au Jour...
  if methode='1' then Result := Round(StrToDate(DateFin)) - Round(StrToDate(DateDeb)); }

  if StrToInt(methode)>1 then Result := DonneNbIntervalleMailleDateDebDateFin(methode,DateDeb,DateFin)
  else Result := 1;

end;

function NbMajCoeff(codesession:hString):integer;
var nivMax:integer;
Q : TQuery;
begin
   if DonneMethodeSession(codeSession) = '1' then
   begin
     nivMax:=ChercheNivMaxSession(codeSession);
     nivMax:=nivMax+1;
     Q:=MOpenSql('SELECT COUNT(QBR_NUMNOEUD)'+
      ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codesession+'" AND'+
      ' (QBR_PRCTVARIATION1<>100 OR QBR_PRCTVARIATION2<>100 OR QBR_PRCTVARIATION3<>100'+
      ' OR QBR_PRCTVARIATION4<>100 OR QBR_PRCTVARIATION5<>100 OR QBR_PRCTVARIATION6<>100'+
      ' OR QBR_PRCTVARIATIONQ<>100 OR QBR_VALMODIF="X") AND QBR_NIVEAU='+IntToSTR(nivMax),'SynScriptBP (NbMajCoeff).',true );
     Result := Q.Fields[0].AsInteger;
     Ferme(Q);
   end else Result := 0;
end;

function FuncVerifDateArbre(parms: array of variant; nb: integer): variant;
var Session:hString;
 DateSelect,DateDeb,DateFin:TDateTime;
 Methode,JourS,JourFS :hString;
 Q: TQUery;
 PremierJourS:integer;
 an,mois,day:word;
 PerDeb : TPeriodeP;
begin
  DateDeb:=idate1900;
  DateFin:=idate1900;
  Session:=hString(parms[0]);
  if Pos(' ',hString(parms[1]))=0 then
  begin
    DateSelect:=StrToDateTime(hString(parms[1]));

    Q :=OpenSQL('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_BPINITIALISE FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',True)  ;
    If Not Q.Eof Then
    begin
      DateDeb:=Q.Fields[0].AsDateTime ;
      DateFin:=Q.Fields[1].AsDateTime ;
      methode:=Q.Fields[2].AsString;
    end;
    Ferme(Q) ;

    if ((DateDeb <= DateSelect) AND (DateSelect <= DateFin)) then
    begin
      //Methode = '0' -> Loi d'Eclatement
      if methode='0' then Result := 0;

      { EVI / Contrôle pour session non initialisée --> 1}
      if methode='' then methode:='1';

      { EVI / Session initialisée au Jour...
      if methode='1' then Result := Round(StrToDate(DateFin)) - Round(StrToDate(DateDeb)); }
      if methode='2' then
      begin
        PremierJourS:=V_PGI.PremierJourSemaine;
        case PremierJourS of
          1 : begin
                JourS:=TraduireMemoire('Dimanche');
                JourFS:=TraduireMemoire('Samedi');
              end;
          2 : begin
                JourS:=TraduireMemoire('Lundi');
                JourFS:=TraduireMemoire('Dimanche');
              end;
          7 : begin
                JourS:=TraduireMemoire('Samedi');
                JourFS:=TraduireMemoire('Vendredi');
              end;
          else JourS:=TraduireMemoire('Lundi');
        end;

        if ((DayOfWeek(DateSelect)<>PremierJourS)) then
        begin
          DateSelect:=DateSelect-(DayOfWeek(DateSelect)-PremierJourS);
          if PGIAsk( Format_( TraduireMemoire( TexteMessage[ 1 ] ), [ JourS ]) + #13#10 +
                     Format_( TraduireMemoire( TexteMessage[ 2 ] ), [ JourS + ' ' + DateTimeToStr(DateSelect) ] ) ) = mrYes
          then Result := DateTimeToStr(DateSelect)
          else Result := -1;
        end
        else Result := 0;
      end;

      if methode='3' then
      begin
        DecodeDate(DateSelect,an,mois,day);
        if ((day<>16) and (day<>1)) then
        begin
          if day>16 then DateSelect:=DateSelect-(day-16) else DateSelect:=DateSelect-(day-1);
          if PGIAsk( TraduireMemoire( TexteMessage[ 3 ] ) + #13#10 +
                     Format_( TraduireMemoire( TexteMessage[ 2 ]), [DateTimeToStr(DateSelect)]) )= mrYes
          then Result := DateTimeToStr(DateSelect)
          else Result := -1;
        end
        else Result := 0;
      end;

      if methode='4' then
      begin
        if (DateSelect<>(DEBUTDEMOIS(DateSelect))) then
        begin
          if PGIAsk( TraduireMemoire( TexteMessage[ 11 ] ) + #13#10 +
             Format_( TraduireMemoire( TexteMessage[ 12 ] ), [ DateTimeToStr( DEBUTDEMOIS( DateSelect ) ) ] ) ) = mrYes
          then Result := DateTimeToStr( DEBUTDEMOIS( DateSelect ) )
          else Result := -1;
        end
        else Result := 0;
      end;

      if methode='5' then { EVI / ToDo : Mois 4-4-5 }
      begin
        PerDeb:=FindPeriodeP(DateDeb);
         if (DateSelect<>(PerDeb.datedeb)) then
        begin
           PGIInfo( 'Pour une session initialisée par mois 4-4-5, la date sélectionnée doit correspondre au premier jour d''une période.');
          Result := -1;
        end
        else Result := 0;
      end;

      if methode='6' then
      begin
        DecodeDate(DateSelect,an,mois,day);
        if (day<>1) and ((mois<>1) or (mois<>4) or (mois<>7) or (mois<>10)) then
        begin
          if ((1 <= mois) AND (mois < 4)) then DateSelect:=EncodeDate(an,01,01);
          if ((4 <= mois) AND (mois < 7)) then DateSelect:=EncodeDate(an,04,01);
          if ((7 <= mois) AND (mois < 10)) then DateSelect:=EncodeDate(an,07,01);
          if (mois >= 10) then DateSelect:=EncodeDate(an,10,01);
          if PGIAsk( TraduireMemoire( TexteMessage[ 13 ] ) + #13#10 + Format_( TraduireMemoire( TexteMessage[ 12 ] ),[ DateTimeToStr( DateSelect ) ] ) ) = mrYes
          then Result := DateTimeToStr(DateSelect)
          else Result := -1;
        end
        else Result := 0;
      end;

      if methode='7' then
      begin
        DecodeDate(DateSelect,an,mois,day);
        if (day<>1) and ((mois<>1) or (mois<>4) or (mois<>7) or (mois<>10)) then
        begin
          if ((1 <= mois) AND (mois < 5)) then DateSelect:=EncodeDate(an,01,01);
          if ((5 <= mois) AND (mois < 9)) then DateSelect:=EncodeDate(an,05,01);
          if (mois >= 9) then DateSelect:=EncodeDate(an,09,01);
          if PGIAsk( TraduireMemoire( TexteMessage[ 14 ] ) + #13#10 + Format_( TraduireMemoire( TexteMessage[ 12 ] ), [ DateTimeToStr( DateSelect ) ] ) ) = mrYes
          then Result := DateTimeToStr( DateSelect )
          else Result := -1;
        end
        else Result := 0;
      end;

    end
    else
    begin
      PGIInfo('La date sélectionnée ne correspond pas à la période définie pour l''objectif.');
      Result := -1;
    end;

  end else Result := -1;
end;


{ EVI / Ajout fonction pour récupérer requête SQL }
function VerifPos( Sens : boolean; Req, Letter : hString; Pos : integer ) : integer;
var
  x, nbchar, max : integer;
  verif : hString;
begin
  x := 0;
  nbchar := length( letter );
  if sens = true then
    max := length( Req ) - Pos
  else
    max := pos;
  while verif <> letter do
  begin
    if sens = true then
      verif := Copy( Req, pos + x, nbchar )
    else
      verif := Copy( Req, pos - x, nbchar );
    x := x + 1;
    if x = max + 1 then
      break;
  end;
  if x = max + 1 then
    Result := -1
  else if sens = true then
    Result := pos + x
  else
    Result := pos - x;
end;

procedure ValidationBP(const codeSession:hString);
begin
 MExecuteSql('UPDATE QBPSESSIONBP SET QBS_SESSIONVALIDE="X",'+
             'QBS_DATEVALIDATION="'+USDATETIME(date)+
             '" WHERE QBS_CODESESSION="'+codeSession+'"',
             'SynScriptBP (ValidationBP).');
end;

procedure SynInitBP;
begin
  // RegisterAglProc ( '...', False, 1, Agl.... ) ;
  // RegisterAglFunc ( '...', False, 1, Agl.... ) ;

  RegisterAglFunc('SynStrToFloat', False, 1, FuncSynStrToFloat);
  RegisterAglFunc('SynGetCtx', False, 1, FuncSynGetCtx);
  RegisterAglFunc('SynWhereCtx', False, 1, FuncSynWhereCtx);
  RegisterAglFunc('SynBPDateFinAnnee', False, 0, FuncSynBPDateFinAnnee);
  RegisterAglFunc('SynBPDateJour', False, 0, FuncSynBPDateJour);
  RegisterAglFunc('SynBPDateDebutAnnee', False, 0, FuncSynBPDateDebutAnnee);
  RegisterAglFunc('SynBPDateFinAnneeMoinsUn', False, 0, FuncSynBPDateFinAnneeMoinsUn);
  RegisterAglFunc('SynBPDateDebutAnneeMoinsUn', False, 0, FuncSynBPDateDebutAnneeMoinsUn);
  RegisterAglFunc('SynDate', False, 1, FuncSynDate);
  RegisterAglFunc('SynQuestionSupp', False, 1, FuncSynQuestionSupp);
  RegisterAglFunc('SynMakeDate', False, 2, FuncSynMakeDate);
  RegisterAglFunc('TrouveArgument', False, 3, FuncTrouveArgument);
  RegisterAglFunc('SynDonneLibBP', False, 6, FuncSynDonneLibBP);
  RegisterAglFunc('SynBPChercheAxes', False, 1, FuncSynBPChercheAxes);
  RegisterAglFunc('SynBPControleValiditeCas4', False, 6, FuncSynBPControleValiditeCas4);
  RegisterAglFunc('SynBPValeurUnique', False, 4, FuncSynBPValeurUnique);
  RegisterAglFunc('SynBPSessionUnique', False, 1, FuncSynBPSessionUnique);
  RegisterAglFunc('SynBPD1infD2', False, 4, FuncSynBPD1infD2);
  RegisterAglFunc('SynBPControleAxes', False, 3, FuncSynBPControleAxes);
  RegisterAglFunc('SynBPIncrementeNumNoeud', False, 1, FuncSynBPIncrementeNumNoeud);
  RegisterAglFunc('SynBPControleCodeAxeDiff', False, 10, FuncSynBPControleCodeAxeDiff);
  RegisterAglFunc('SynBPControleValeurAxeDiff', False, 3, FuncSynBPControleValeurAxeDiff);
  RegisterAglFunc('SynBPDonneLibelleCodeAxe', False, 2, FuncSynBPDonneLibelleCodeAxe);          //PIVOT
  RegisterAglFunc('SynBPDonneCodeAxe', False, 2, FuncSynBPDonneCodeAxe);                        //PIVOT
  RegisterAglFunc('SynBPDonneLibelleValeurAxe', False, 2, FuncSynBPDonneLibelleValeurAxe);      //PIVOT
  RegisterAglFunc('SynBPDonneSessionOk', False, 1, FuncSynBPDonneSessionOk);
  RegisterAglFunc('SynBPDonneCodeTablette', False, 1, FuncSynBPDonneCodeTablette);
  RegisterAglFunc('SynBPLanceChxValeurAxe', False, 2, FuncSynBPLanceChxValeurAxe);
  RegisterAglProc('SynBPInitialise', False, 1, ProcSynBPInitialise);
  RegisterAglFunc('SynBPInitialiseDelaiOk', False, 1, FuncSynBPInitialiseDelaiOk);
  RegisterAglFunc('SynBPPGIOrli', False, 0, FuncSynBPPGIOrli);
  RegisterAglFunc('SynBPOkOrli', False, 0, FuncSynBPOkOrli);
  RegisterAglFunc('SynBPOkPgiEnt', False, 0, FuncSynBPOkPgiEnt);
  RegisterAglFunc('SynDonneLibCodeAxeBP', False, 1, FuncSynDonneLibCodeAxeBP);
  RegisterAglFunc('SynBPValeurAxePlus', False, 1, FuncSynBPValeurAxePlus);
  RegisterAglFunc('SynBPControleValiditeCodeAxe', False, 3, FuncSynBPControleValiditeCodeAxe);
  RegisterAglFunc('SynBPControleValiditeValeurAxe', False, 3, FuncSynBPControleValiditeValeurAxe);
  RegisterAglFunc('SynBPTabletteCodeAxe', False, 0, FuncSynBPTabletteCodeAxe);
  RegisterAglFunc('SynBPTabletteValeurAxe', False, 0, FuncSynBPTabletteValeurAxe);
  RegisterAglFunc('SynBPTabletteValeurAxeSaiCmd', False, 0, FuncSynBPTabletteValeurAxeSaiCmd);
  RegisterAglFunc('SynBPIncrementeSession', False, 1, FuncSynBPIncrementeSession);
  RegisterAglProc('SynBPValidationVersSCM', False, 1, ProcSynBPValidationVersSCM);
  RegisterAglFunc('SynBPDonneLibelleSession', False, 2, FuncSynBPDonneLibelleSession);
  RegisterAglFunc('SynBPArbreControleValeurAxe', False, 4, FuncSynBPArbreControleValeurAxe);
  RegisterAglFunc('SynBPArbreControleDate', False, 4, FuncSynBPArbreControleDate);
  RegisterAglFunc('SynBPTablettePlus', False, 1, FuncSynBPTablettePlus);
  RegisterAglFunc('SynBPTabletteValeurAxeC', False, 1, FuncSynBPTabletteValeurAxeC);
  RegisterAglFunc('SynBPTabPlusValeurAxeSaiCmd', False, 0, FuncSynBPTabPlusValeurAxeSaiCmd);
  RegisterAglFunc('SynBPPlusValeurAxe', False, 1, FuncSynBPPlusValeurAxe);
  RegisterAglFunc('SynBPDonneDateDebSaison', False, 1, FuncSynBPDonneDateDebSaison);
  RegisterAglFunc('SynBPDonneDateFinSaison', False, 1, FuncSynBPDonneDateFinSaison);
  RegisterAglFunc('SynBPDonneDateDebSession', False, 1, FuncSynBPDonneDateDebSession);
  RegisterAglFunc('SynBPDonneDateFinSession', False, 1, FuncSynBPDonneDateFinSession);
  RegisterAglProc('SynBPMAJAxeSaison', False, 2, ProcSynBPMAJAxeSaison);
  RegisterAglFunc('SynBPTablettePlusSession', False, 1, FuncSynBPTablettePlusSession);
  RegisterAglFunc('SynBPSessionNivMax', False, 1, FuncSynBPSessionNivMax);
  RegisterAglFunc('SynBPDonneMethodeSession', False, 1, FuncSynBPDonneMethodeSession);
  RegisterAglProc('SynBPExportVersOrliPap', False, 1, ProcSynBPExportVersOrliPap);
  RegisterAglProc('SynBPExportVersOrliWeb', False, 1, ProcSynBPExportVersOrliWeb);
  RegisterAglFunc('SynBPSessionCodeAxe', False, 1, FuncSynBPSessionCodeAxe);
  RegisterAglFunc('SynBPSessionBudgetOk', False, 1, FuncSynBPSessionBudgetOk);
  RegisterAglFunc('SynBPSessionPermanent', False, 1, FuncSynBPSessionPermanent);
  RegisterAglFunc('SynBPDonneTabVisible', False, 1, FuncSynBPDonneTabVisible);
  RegisterAglFunc('SynChpObligatoire', False, 2, FuncSynChpObligatoire);
  RegisterAglFunc('SynBPBtnInitDelaiObj', False, 1, FuncSynBPBtnInitDelaiObj);
  RegisterAglFunc('SynBPBtnInitGlobalObj', False, 1, FuncSynBPBtnInitGlobalObj);
  RegisterAglFunc('SynBPBtnInitCoeff', False, 1, FuncSynBPBtnInitCoeff);
  RegisterAglFunc('SynBPBtnInitPrev', False, 1, FuncSynBPBtnInitPrev);
  RegisterAglFunc('SynBPBtnInitPrevDelai', False, 1, FuncSynBPBtnInitPrevDelai);
  RegisterAglFunc('SynBPBtnInitPrevTaille', False, 1, FuncSynBPBtnInitPrevTaille);
  RegisterAglFunc('SynBPSessionValide', False, 1, FuncSynBPSessionValide);
  RegisterAglFunc('SynBPLanceCalculInitObjectif', False, 8, FuncSynBPLanceCalculInitObjectif);
  RegisterAglProc('SynBPLanceCalculInitCoeff', False, 0, ProcSynBPLanceCalculInitCoeff);
  RegisterAglProc('SynBPLanceCalculInitPrevision', False, 1, ProcSynBPLanceCalculInitPrevision);
  RegisterAglProc('SynBPLanceCalculEclatementLoi', False, 2, ProcSynBPLanceCalculEclatementLoi);
  RegisterAglProc('SynBPLanceValidation', False, 1, ProcSynBPLanceValidation);
  RegisterAglProc('SynBPLanceDeValidation', False, 1, ProcSynBPLanceDeValidation);
  RegisterAglProc('SynBPMiseAJourTabletteMaille', False, 2, ProcSynBPMiseAJourTabletteMaille);
  RegisterAglProc('SynBPModifSessionVue', False, 2, ProcSynBPModifSessionVue);
  RegisterAglFunc('SynBPIncrementeNumSession', False, 0, FuncSynBPIncrementeNumSession);
  RegisterAglFunc('SynBPValRestrictionValaxe1', False, 1, FuncSynBPValRestrictionValaxe1);
  RegisterAglFunc('SynBPValRestrictionValaxe2', False, 1, FuncSynBPValRestrictionValaxe2);
  RegisterAglFunc('SynBPValRestrictionValaxe3', False, 1, FuncSynBPValRestrictionValaxe3);
  RegisterAglFunc('SynBPValRestrictionValaxe4', False, 1, FuncSynBPValRestrictionValaxe4);
  RegisterAglFunc('SynBPValaxe1234ValRestriction', False, 4, FuncSynBPValaxe1234ValRestriction);
  RegisterAglFunc('SynBPOkListeSem', False, 0, FuncSynBPOkListeSem);
  RegisterAglFunc('SynBPOkListeQuinz', False, 0, FuncSynBPOkListeQuinz);
  RegisterAglFunc('SynBPOkListeMois', False, 0, FuncSynBPOkListeMois);
  RegisterAglFunc('SynBPOkListeMois445', False, 0, FuncSynBPOkListeMois445);
  RegisterAglFunc('SynBPOkListeTrim', False, 0, FuncSynBPOkListeTrim);
  RegisterAglFunc('SynBPOkListeQuadri', False, 0, FuncSynBPOkListeQuadri);
  RegisterAglProc('SynBPMAJDatesLoi', False, 3, ProcSynBPMAJDatesLoi);
  RegisterAglFunc('SynBPLibelleMaille', False, 1, FuncSynBPLibelleMaille);
  RegisterAglFunc('SynBBOkLoi', False, 1, FuncSynBBOkLoi);
  RegisterAglFunc('SynBBHShowmessage', False, 1, FuncSynBBHShowmessage);
  RegisterAglProc('SynBPDuplicationArbre', False, 2, ProcSynBPDuplicationArbre);
  RegisterAglProc('SynBPDuplicationArbrePartielle', False, 2, ProcSynBPDuplicationArbrePartielle);
  RegisterAglProc('SynBPDuplicationArbreBP', False, 2, ProcSynBPDuplicationArbreBP);
  RegisterAglProc('SynBPDuplicationArbreBPPartielle', False, 2, ProcSynBPDuplicationArbreBPPartielle);
  RegisterAglFunc('SynGestTaille', False, 0, FuncSynGestTaille);
  RegisterAglFunc('VerifDateArbre', False, 1, FuncVerifDateArbre);
  RegisterAglFunc('SynBPExportSessionValide', False, 0, FuncSynBPExportSessionValide);
  RegisterAglFunc('SynBPOkDelai', False, 1, FuncSynBPOkDelai);
  RegisterAglFunc('SynBPOkTaille', False, 1, FuncSynBPOkTaille);
  RegisterAglProc('SynBPSessionMAJ', False, 2, ProcSynBPSessionMAJ);
  RegisterAglFunc('SynBPSessionDelete', False, 1, FuncSynBPSessionDelete); // 15727

end;

initialization
  SynInitBP;

end.
