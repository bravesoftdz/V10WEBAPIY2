{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit MZSUtil;

interface

uses Forms,IniFiles,dialogs,SysUtils,TabNotBk,Controls,buttons,
     Classes,HCtrls,HCompte,ComCtrls,Hstatus,
     stdctrls,Grids,HQuickRP, ExtCtrls, Windows,
     registry,HQry,WinProcs,HDebug,Ent1,
     {$IFNDEF IMP}
     ModifZS,
     {$ENDIF}
     ParamSoc,
     hmsgbox, { Ajout ME le 11/07/2001 hmsgbox}
     uTob,
     {$IFDEF EAGLCLIENT}

     {$ELSE}
       {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
       HDB,
       DB,
       DBGrids,
       DBCtrls,
     {$ENDIF EAGLCLIENT}
     {$IFNDEF CCMP}
     {$IFNDEF IMP}
       {$IFNDEF CCSTD}
       {$IFNDEF COMSX}
         BUDGENE_TOM,
         BUDSECT_TOM,
         BUDJAL_TOM,
       {$ENDIF}
       {$ENDIF CCSTD}
       {$ENDIF IMP}
     {$ENDIF CCMP}
     CPGeneraux_TOM,
     CPTiers_TOM,
     CPSection_TOM ,
     ULibanalytique
//     CPMODRESTANA_TOM       {FP 09/01/2006}
     ;

function DoSQLUpdateOnTiersCompl(SQL, Where : String) : string;
{$IFDEF EAGLCLIENT}
function ModifieEnSerie(Table,Faxe : string; G : THGrid; Q : THQuery) : Boolean;
{$ELSE}
function ModifieEnSerie(Table,Faxe : string; G : THDBGrid; Q : THQuery) : Boolean;
{$ENDIF}

// GCO - 04/07/2005
procedure ModifieEnSerieEcriture( vTobListeEcr : Tob );
// FIN GCO

function SQLUpdate(StSQL,Ctrl,Table : string): string;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HEnt1;

type
  TErrorItem = array[1..MaxAxe] of record
    Nb:    Integer;     {Nb de fois qu'une erreur est rencontrée}
    Force: Boolean;     {True pour ignorer l'erreur}
  end;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 24/05/2006
Description .. : GCO - 24/05/2006 - FQ 18157
Mots clefs ... :
*****************************************************************}
function DoSQLUpdateOnTiersCompl(SQL, Where : String) : string;
var sz, szChamp, szValeur, szSet, szWhere : String;
//  szAccelerateur : string;  //fb 17/11/2005 FQ 14934
  i : integer;

  lTob : Tob;
  lSTAuxiliaire : string;
  lStTiers      : string;

  function IsTiers(Champ : String) : Boolean;
  begin
    Result := (Pos('T_',Champ)>0);
  end;

BEGIN
  Result := '';
  //szAccelerateur := '';  //fb 17/11/2005 FQ 14934

  i := Pos('"',Where);

  lStAuxiliaire := Trim(Copy(Where,i+1,Length(Where)-i-1));

  szWhere := ' WHERE YTC_AUXILIAIRE = "'+ lStAuxiliaire + '"';

  // GCO - 24/05/2006 - FQ 18157
  if not ExisteSQL('SELECT YTC_AUXILIAIRE FROM TIERSCOMPL ' + szWhere) then
  begin // Création dans TIERSCOMPL vu qu'il n'existe pas.
    lStTiers := GetColonneSQL('TIERS', 'T_TIERS', 'T_AUXILIAIRE = "' + lStAuxiliaire + '"');
    lTob := Tob.Create('TIERSCOMPL', nil, -1);
    lTob.InitValeurs;
    lTob.PutValue('YTC_AUXILIAIRE', lStAuxiliaire);
    lTob.PutValue('YTC_TIERS', lStTiers);
    lTob.InsertDB(nil);
    FreeAndNil(lTob);
  end;

  while SQL <> '' do
  begin
    sz := ReadTokenSt(SQL);
    i := Pos('=',sz);
    if i<=0 then Continue;
    szChamp:=Trim(Copy(sz,1,i-1));
{b fb 17/11/2005 FQ 14934}
{    if szChamp<>'YTC_ACCELERATEUR' then  // On supprime le champ accélérateur de la liste pour le mettre à jour séparémment
    begin  }
{e fb 17/11/2005 FQ 14934}
      if ISTiers(szChamp) then begin
        {b FP 13/10/2005 FQ16855 Il faut remetre le séparateur de chaine car
           la function SQLUpdate utilise ReadTokenSt}
        if (Result <> '') and (sz <> '') then
          Result := Result+';';
        {e FP 13/10/2005 FQ16855}
        Result := Result+sz;
        if (SZChamp='T_DATEMODIF') then szChamp := 'YTC_DATEMODIF'
                                   else continue;
      end;
      szValeur:=Trim(Copy(sz,i+1,Length(sz)-i));
      szSet := szSet + szChamp + '=' + szValeur;
{b fb 17/11/2005 FQ 14934}
{    end else
    begin
      szAccelerateur := szChamp + '=' + Trim(Copy(sz,i+1,Length(sz)-i-1));
    end;}
{e fb 17/11/2005 FQ 14934}
  end;
  if szSet<>'' then begin
    //BeginTrans; //fb 17/11/2005 FQ 14934
    ExecuteSQL('UPDATE TIERSCOMPL SET ' + szSet + szWhere);
{b fb 17/11/2005 FQ 14934}
    {if (szAccelerateur <>'') then ExecuteSQL('UPDATE TIERSCOMPL SET '+szAccelerateur+szWhere+' AND YTC_SCHEMAGEN<>""');
    CommitTrans; }
{e fb 17/11/2005 FQ 14934}
  end;
end;
{b FP 29/12/2005}
////////////////////////////////////////////////////////////////////////////////


procedure AnalyseSQLStatement(SQL: String; LstChamp: HTStrings; LstVal: HTStrings);
var
  sz, szChamp, szValeur: String;
  i :       integer;
begin
  LstChamp.Clear;
  LstVal.Clear;
  While SQL<>'' do
    begin
    sz := ReadTokenSt(SQL);
    i := Pos('=',sz);
    if i<=0 then Continue;
    szChamp:=Trim(Copy(sz,1,i-1));
    szValeur:=Trim(Copy(sz,i+1,Length(sz)-i));
    LstChamp.Add(szChamp);
    LstVal.Add(szValeur);
    end;
end;

function VerifModelRestriction(OkRaf: boolean; szCompte: String; QModel: TQuery; LstChamp, LstVal: HTStrings; var ErrorMsg: array of TErrorItem): Boolean;
  function VerifModeleVentil(Modele, Axe: String): Boolean;
  var
    NumAxe:   String;
    SQL:      String;
    Restriction: TRestrictionAnalytique;
  begin
    Restriction := TRestrictionAnalytique.Create;
    try
      Result := True;
      Modele := Trim(Modele);
      if Modele <> '' then
        begin
        NumAxe :=  Axe[Length(Axe)];
        SQL    := 'SELECT V_SECTION' +
                  ' FROM VENTIL' +
                  ' WHERE V_COMPTE="'+szCompte+'"'+
                  '  AND V_NATURE="GE'+NumAxe+'"' +
                  '  AND '+Restriction.GetClauseModeleCompteInterdit(Modele, Axe, 'VENTIL');
        Result := not ExisteSQL(SQL);
        end;
    finally
      Restriction.Free;
      end;
  end;
var
  i:        Integer;
  Index:    Integer;
  Modele:   String;
  MsgError: String;
  Axe:      String;
  iAxe:     Integer;
  Delete:   String;           {Liste des axes dont il ne faut pas faire la MAJ}
begin           {VerifModelRestriction}
  Result := True;
  Delete := '';
  for i:=0 to LstChamp.Count-1 do
    begin
    iAxe   := StrToInt(LstChamp[i][Length(LstChamp[i])]);
    Axe    := 'A'+IntToStr(iAxe);
    Modele := StringReplace(LstVal[i], '"', '', [rfReplaceAll, rfIgnoreCase]);
    Modele := Trim(StringReplace(Modele, ',', '', [rfReplaceAll, rfIgnoreCase]));

    if (not QModel.EOF) and (QModel.FindField(LstChamp[i]).AsString<>'') and
       (QModel.FindField(LstChamp[i]).AsString<>Modele) then
      begin
      if OkRaf then
        MsgError := 'Des comptes généraux ont déjà un modèle associé pour l''axe '+
          IntToStr(iAxe)+'.'+Chr($D)+Chr($A)+'Voulez-vous les écraser ?'
      else
        MsgError := 'Le compte général '+szCompte+' a déjà un modèle associé pour l''axe '+
          IntToStr(iAxe)+'.'+Chr($D)+Chr($A)+'Voulez-vous l''écraser ?';
      if (not OkRaf) then
        begin
        ErrorMsg[0,iAxe].Force := HShowMessage('0;Restriction analytique;'+MsgError+';Q;YN;N;N;','','')=mrYes
        end
      else if ErrorMsg[0,iAxe].Nb = 0 then
        begin
        ErrorMsg[0,iAxe].Force := HShowMessage('0;Restriction analytique;'+MsgError+';Q;YN;N;N;','','')=mrYes;
        Inc(ErrorMsg[0,iAxe].Nb);
        end;
      if not ErrorMsg[0,iAxe].Force then
        Delete := Delete+IntToStr(iAxe);
      end;
    if (Pos(IntToStr(iAxe), Delete)=0) and (not VerifModeleVentil(Modele, Axe)) then
      begin
      {b FP 19/04/2006 FQ17831}
      if OkRaf then
        MsgError := 'Pour l''axe '+IntToStr(iAxe)+', il existe des incomptabilités entre la ventilation par défaut'+
         ' des comptes et'+Chr($D)+Chr($A)+' les modèles de restriction à rattacher aux comptes généraux.'+Chr($D)+Chr($A)+
         ' La mise à jour de ces comptes ne s''est pas effectuée.'
      else
        MsgError := 'Pour l''axe '+IntToStr(iAxe)+', il existe des incomptabilités entre la ventilation par défaut'+
         ' du compte général '+szCompte+' et'+Chr($D)+Chr($A)+' le modèle de restriction à rattacher.'+Chr($D)+Chr($A)+
         ' La mise à jour de ce compte ne s''est pas effectuée.';
      {e FP 19/04/2006}
      if (not OkRaf) or (ErrorMsg[1,iAxe].Nb = 0) then
        begin
        PGIError(MsgError, 'Restriction analytique');
        Inc(ErrorMsg[1,iAxe].Nb);
        end;
      Delete := Delete+IntToStr(iAxe);
      Result := False;         {FP 19/04/2006 FQ17832}
      end;
    end;
  for i:=1 to Length(Delete) do
    begin
    Index := LstChamp.IndexOf('CLA_CODE'+Delete[i]);
    if Index <> -1 then
      begin
      LstChamp.Delete(Index);
      LstVal.Delete(Index);
      end;
    end;
end;

function DoSQLUpdateOnRestrictAna(SQL, Where : String; OkRaf: Boolean; var ErrorMsg: array of TErrorItem) : string;
  function IsGeneral(Champ : String) : Boolean;
  begin
    Result := (Pos('G_',Champ)>0);
  end;
var
  sz, szChamp, szSet, szWhere : String;
  szCompte: String;
  i :       integer;
  Insert:   Boolean;
  Q:        TQuery;
  LstChamp: HTStrings;
  LstVal:   HTStrings;
  Ok:       Boolean;        {FP 19/04/2006 FQ17832}
BEGIN
  Result   := '';
  LstChamp := HTStringList.Create;
  LstVal   := HTStringList.Create;
  try

  // Clause Where
  i := Pos('=',Where);
  szCompte := Trim(StringReplace(Copy(Where,i+1,Length(Where)-i), '"', '', [rfReplaceAll,rfIgnoreCase]));
  szWhere  := ' WHERE CLA_GENERAL="'+szCompte+'"';

  AnalyseSQLStatement(SQL, LstChamp, LstVal);
  for i:=0 to LstChamp.Count-1 do
    begin
    sz := LstChamp[i]+'='+LstVal[i];
    szChamp := LstChamp[i];
    if IsGeneral(szChamp) then
      begin
      if (Result <> '') and (sz <> '') then
        Result := Result+';';
      Result := Result+sz;
      end
    end;

  for i:=LstChamp.Count-1 downto 0 do {Retire les champs n'appartenant pas à la table CLIENGENEMODELA}
    if Pos(TableToPrefixe('CLIENGENEMODELA'),LstChamp[i]) = 0 then
      begin
      LstChamp.Delete(i);
      LstVal.Delete(i);
      end;

  if LstChamp.Count>0 then
    begin
    if Pos(';', Result) = 0 then   {Il ne reste que le champ G_DATEMODIF}
      Result := '';

    Q  := OpenSQL('SELECT * from CLIENGENEMODELA'+szWhere, True);
    try
      Insert := Q.EOF;
      Ok := VerifModelRestriction(OkRaf, szCompte, Q, LstChamp, LstVal, ErrorMsg);                 {FP 19/04/2006 FQ17832}
      szSet := '';
      for i:=0 to LstChamp.Count-1 do
        szSet := szSet + LstChamp[i]+'='+LstVal[i];
      szSet := Trim(szSet);
      if szSet <> '' then
        begin
        Delete(szSet, Length(szSet), 1); {Supprime la dernière virgule}
        BeginTrans;
        try
          if Insert then
            ExecuteSQL('INSERT INTO CLIENGENEMODELA (CLA_GENERAL) values ("'+szCompte+'")');
          ExecuteSQL('UPDATE CLIENGENEMODELA SET ' + szSet + szWhere);
          CommitTrans;
        except
          Rollback;
          raise;
          end;
        end;
      {b FP 19/04/2006 FQ17832: En cas d'erreur force l'affichage de la fiche compte}
      if (not OkRaf) and (not ok) then
        begin
        if Result = '' then
          Result := 'G_GENERAL="'+szCompte+'"'
        else
          Result := 'G_GENERAL="'+szCompte+'",;'+Result;
        end;
      {e FP 19/04/2006}
    finally
      Ferme(Q);              {FP 19/04/2006}
      FreeAndNil(Q);
      end;
    end;
  finally
    LstChamp.Free;
    LstVal.Free;
    end;
end;
{e FP 29/12/2005}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/08/2006
Modifié le ... :   /  /    
Description .. : - LG - 03/08/2006 - on bloque la mise à jour en serie sur les 
Suite ........ : axe de tva
Mots clefs ... : 
*****************************************************************}
procedure ModifieLe(OkRaf: boolean ; Q : TQuery ; FCode,FCtrl,StSQL,Table,FAxe : String; var ErrorMsg: array of TErrorItem) ;
var StWhere,StCtrl,SQL : String ;
    FF                 : TField ;
    St2                : string;
    lStA               : string ;
    lStC               : string ;
BEGIN
{b FP 29/12/2005}
  StWhere:=' Where '+ FCode+'="'+ Q.FindField(FCode).AsString+'"';

  if ( Table = 'GENERAUX' ) and GetParamSocSecur('SO_CPPCLSAISIETVA',false) then
   begin
    lStA := 'G_VENTILABLE' + Copy(GetParamSocSecur('SO_CPPCLAXETVA', '  '),2,1) ;
    lStC :=  Copy(Q.FindField(FCode).AsString,1,2) ;
    if VH^.OuiTvaEnc then
     begin
      if ( lStC = '41' ) and ( Pos( lStA  , stSQL)  > 0 ) then exit
     end
      else
       if ( lStC = '70' ) and ( Pos( lStA  , stSQL)  > 0  ) then exit ;
   end ;

  FF:=Q.FindField(FCtrl) ;
  if ((FCtrl<>'') and (FF<>Nil)) then
    StCtrl:=Q.FindField(FCtrl).AsString
  else
    StCtrl:='';

  if (Table='GENERAUX') then
    StSQL := DoSQLUpdateOnRestrictAna(StSQL,StWhere,OkRaf, ErrorMsg);
{e FP 29/12/2005}

  if OkRaf then
   BEGIN
   {FP 29/12/2005
   StWhere:=' Where '+ FCode+'="'+ Q.FindField(FCode).AsString+'"';
   FF:=Q.FindField(FCtrl) ;
   if ((FCtrl<>'') and (FF<>Nil)) then StCtrl:=Q.FindField(FCtrl).AsString else StCtrl:='';}
   if (Table='TIERS') then StSQL := DoSQLUpdateOnTiersCompl(StSQL,StWhere); // 12399
   SQL:=SQLUpdate(StSQL,StCtrl,Table) ;
   if SQL<>'' then BEGIN BeginTrans; ExecuteSQL(SQL+StWhere); CommitTrans; END ;
   END else
   BEGIN
   if Table='GENERAUX' then
   begin

        if (ctxPCL in V_PGI.PGIContexte) then
        begin
          St2 := StSQL;
          St2:=ReadTokenPipe(St2, ',');
          if (St2 = 'G_VENTILABLE1="-" ') or (St2 = 'G_VENTILABLE2="-" ')then
             if HShowMessage('0;Suppression analytiques;'+'Confimez-vous la suppression de l''ensemble des données du compte : '+Q.FindField(FCode).AsString+';Q;YN;N;N;','','')<>mrYes then
                  exit;
        end;

        if StSQL <> '' then    {FP 29/12/2005}
          FicheGeneMZS('',Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL)
   end else
      if Table='TIERS' then FicheTiersMZS(Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL) else
         if Table='SECTION' then FicheSectionMZS(FAxe,Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL)else
{$IFNDEF CCMP}
{$IFNDEF CCSTD}
{$IFNDEF COMSX}
{$IFNDEF IMP}
            if Table='BUDGENE' then FicheBudGeneMZS(Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL) else
               if Table='ARTICLE' then {FicheArticleMZS('',Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL)} ;
                  if Table='BUDSECT' then FicheBudsectMZS(Faxe,Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL) else
                      if Table='BUDJAL' then FicheBudjalMZS(Faxe,Q.FindField(FCode).AsString,taModifEnSerie,0,StSQL) else
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
                       ;

   END;
{$IFDEF COMPTA}
   // ajout me historisation des modifications pour les synchros
  if ((Table='GENERAUX') or (Table='TIERS') or (Table='SECTION')) then
        MAJHistoparam (Table,  Q.FindField(FCode).AsString);
{$ENDIF}

END ;

{$IFDEF EAGLCLIENT}
function ModifieEnSerie(Table,Faxe : string; G : THGrid; Q : THQuery) : Boolean;
{$ELSE}
function ModifieEnSerie(Table,Faxe : string; G : THDBGrid; Q : THQuery) : Boolean;
{$ENDIF}
var OkRaf                    : boolean;
    StSQL : string;
    FCode,FCtrl         : string;
    i                        : integer;
{$IFNDEF EAGLCLIENT}
    BM                       : TBookmark;
{$ENDIF}
    Pref                     : String ;
    Ladate                   : String ;
    LadateRaf                : String ;
    St2                      : string;
    ErrorMsg                 : array[0..1] of TErrorItem;   {FP 29/12/2005}
BEGIN
FCode:='' ; Result := True;
FillChar(ErrorMsg, sizeof(ErrorMsg), #0);                   {FP 29/12/2005} 
if Table='GENERAUX' then
   BEGIN FCode:='G_GENERAL' ; FCtrl:='G_NATUREGENE' ; FAxe:='' ; Pref:='G_' ; END else
if Table='TIERS' then
   BEGIN FCode:='T_AUXILIAIRE' ; FCtrl:='T_NATUREAUXI' ; FAxe:='' ; Pref:='T_' ; END else
if Table='SECTION' then
   BEGIN FCode:='S_SECTION' ; FCtrl:='' ; FAxe:=Faxe  ; Pref:='S_' ; END else
if Table='BUDGENE' then
   BEGIN FCode:='BG_BUDGENE' ; FCtrl:='' ; FAxe:=''  ; Pref:='BG_' ; END else
if Table='BUDSECT' then
   BEGIN FCode:='BS_BUDSECT' ; FCtrl:='' ; FAxe:=Faxe  ; Pref:='BS_' ; END else
if Table='ARTICLE' then
   BEGIN FCode:='A_ARTICLE' ; FCtrl:='' ; FAxe:=''  ; Pref:='A_' ; END else
if Table='BUDJAL' then
   BEGIN FCode:='BJ_BUDJAL' ; FCtrl:='' ; FAxe:=Faxe  ; Pref:='BJ_' ; END ;
if FCode='' then Exit ;
OkRaf:=true;
Ladate:=',;'+Pref+'DATEMODIF="'+DateToStr(Date)+'"';
LadateRaf:=',;'+Pref+'DATEMODIF="'+UsDateTime(Date)+'"';
{$IFNDEF IMP}
StSQL:=ModifZoneSerie(Table,Faxe, OkRaf);
if StSQL='' then begin
  Result := False;
  exit ;
end;
{$ENDIF}
if OkRaf then StSQL:=StSql+LaDateRaf
         else StSQL:=StSql+LaDate ;
{$IFNDEF EAGLCLIENT}
BM:=Q.GetBookmark;
{$ENDIF}
  if G.AllSelected then
  begin
    Q.DisableControls ;
    Q.First ;
    // Ajout ME le 12/09/2001
    if (ctxPCL in V_PGI.PGIContexte) and (Table = 'GENERAUX') then
    begin
      St2 := StSQL;
      St2:=ReadTokenPipe(St2, ',');
      if (St2 = 'G_VENTILABLE1="-" ') or (St2 = 'G_VENTILABLE2="-" ')then
         if HShowMessage('0;Suppression analytiques;'+'Confimez-vous la suppression de l''ensemble des données analytiques des comptes selectionnés ?'+';Q;YN;N;N;','','')<>mrYes then
              exit;
    end;

  {$IFDEF EAGLCLIENT}
    while not Q.TQ.Eof do
    begin
      ModifieLe(OkRaf,Q.TQ,FCode,FCtrl,StSQL,Table,FAxe, ErrorMsg);
      Q.TQ.Next;
    end;
  {$ELSE}
    while not Q.Eof do
    begin
      ModifieLe(OkRaf,Q,FCode,FCtrl,StSQL,Table,FAxe, ErrorMsg) ;
      Q.Next ;
    end;
  {$ENDIF}
    Q.EnableControls ;
  end
  else // Pas tous Selected
  begin
    for i:=0 to G.NbSelected-1 do
    begin
      G.GotoLeBookMark(i) ;
  {$IFDEF EAGLCLIENT}
     Q.TQ.Seek(G.Row - 1);
     ModifieLe(OkRaf,Q.TQ,FCode,FCtrl,StSQL,Table,FAxe, ErrorMsg) ;
  {$ELSE}
     ModifieLe(OkRaf,Q,FCode,FCtrl,StSQL,Table,FAxe, ErrorMsg) ;
  {$ENDIF}
    end;
  end;

{$IFNDEF EAGLCLIENT}
 Q.GotoBookmark(BM);
 Q.FreeBookmark(BM);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure ModifieEnSerieEcriture( vTobListeEcr : Tob );
var lBoRafale : Boolean;
    lStUpdateSql : string;
    lStWhere     : string;
    i : integer;
begin
  if vTobListeEcr = nil then Exit;

  // Toujours en Rafale
  lBoRafale := True;
{$IFNDEF IMP}
  lStUpdateSql := ModifZoneSerie('ECRITURE', '', lBoRafale);
{$ENDIF}
  if lStUpdateSql = '' then
  begin
    Exit ;
  end;

  lStUpdateSql := lStUpdateSql + ',;E_DATEMODIF = "' + UsTime(Now) + '"'; // E_DATEMODIF
  lStUpdateSql := lStUpdateSql + ',;E_IO = "X"';                          // E_IO
  lStUpdateSql := SQLUpdate(lStUpdateSql, '', 'ECRITURE') ;

  BeginTrans;
  try
    if lStUpdateSql <> '' then
    begin
      for i := 0 to vTobListeEcr.Detail.Count - 1 do
      begin
        lStWhere := ' WHERE ' +
                    'E_JOURNAL = "' + vTobListeEcr.Detail[i].GetString('E_JOURNAL') + '" AND ' +
                    'E_EXERCICE = "' + vTobListeEcr.Detail[i].GetString('E_EXERCICE') + '" AND ' +
                    'E_DATECOMPTABLE = "' + UsDateTime(vTobListeEcr.Detail[i].GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
                    'E_NUMEROPIECE = ' + IntToStr(vTobListeEcr.Detail[i].GetInteger('E_NUMEROPIECE')) + ' AND ' +
                    'E_NUMLIGNE = ' + IntToStr(vTobListeEcr.Detail[i].GetInteger('E_NUMLIGNE')) + ' AND ' +
                    'E_NUMECHE = ' + IntToStr(vTobListeEcr.Detail[i].GetInteger('E_NUMECHE')) + ' AND ' +
                    'E_QUALIFPIECE = "' + vTobListeEcr.Detail[i].GetValue('E_QUALIFPIECE') + '" AND ' +
                    'E_DATEMODIF = "' + UsTime(vTobListeEcr.Detail[i].GetDateTime('E_DATEMODIF')) + '"';

        if ExecuteSQL( lStUpdateSql + lStWhere ) <> 1 then
        begin
          Exception.Create('La ligne suivante : '+#10#13+
            'Exercice : '+ vTobListeEcr.Detail[i].GetString('E_EXERCICE') +
            ' journal : '+ vTobListeEcr.Detail[i].GetString('E_JOURNAL') + #10#13 +
            'Bordereau n° '+ IntToStr(vTobListeEcr.Detail[i].GetInteger('E_NUMEROPIECE')) +
            ' du ' + DateToStr(vTobListeEcr.Detail[i].GetDateTime('E_DATECOMPTABLE')) +
            ' libellé ' + vTobListeEcr.Detail[i].GetString('E_LIBELLE') + #10#13 +
            ' n''a pas pu être mise à jour.');
        end;
      end;
      CommitTrans;
    end;
  except
    on E: Exception do
    begin
      RollBack;
      PgiError('Erreur, ' + E.Message, 'Fonction : ModifieEnSerieEcriture');
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

(*======================================================================*)
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 21/09/2005
Modifié le ... :   /  /    
Description .. : - LG - 21/09/2005 - Fb 16664 - affectation en auto de la 
Suite ........ : periode si elle n'etait pas renseignée
Mots clefs ... : 
*****************************************************************}
function SQLUpdate(StSQL,Ctrl,Table : string): string;
var St,Champ,Valeur,LeUpdate,LeSet : String;
    i                              : integer;
//    St2                            : string;
{b FP 13/10/2005 FQ16855}
    s, ValeurReel: String;
{e FP 13/10/2005 FQ16855}
BEGIN
LeUpdate:='Update '+Table+' Set ';
Result:='' ; LeSet:='' ;

// gestion des comptes de cutoff
if ( Pos('G_CUTOFF',StSQL) > 0 ) and ( Pos('G_CUTOFFPERIODE',StSQL) = 0 ) then
 StSQL := StSQL + ',;G_CUTOFFPERIODE="CN"' ;


{ Ajout ME le 11/07/2001 }
{if (ctxPCL in V_PGI.PGIContexte) and (Table = 'GENERAUX') then
begin
  St2 := StSQL;
  St2:=ReadTokenPipe(St2, ',');
  if (St2 = 'G_VENTILABLE1="-" ') or (St2 = 'G_VENTILABLE2="-" ')then
     if HShowMessage('0;Suppression analytiques;'+'Confimez-vous la suppression de l''ensemble des données analytiques des comptes selectionnés ?'+';Q;YN;N;N;','','')<>mrYes then
          exit;
end;}
{ fin Ajout ME le 11/07/2001 }

While StSQL<>'' do
  BEGIN
  St:=ReadTokenSt(StSQL);
  i:=Pos('=',St); if i<=0 then Continue;
  Champ:=Trim(Copy(St,1,i-1)); Valeur:=Trim(Copy(St,i+1,Length(St)));
  if Table='GENERAUX' then
     BEGIN
     //SG6 03.03.05
     if (Champ='G_VENTILABLE') then
     begin
       if Valeur = '"X" ,' then
       begin
         for i := 1 to MaxAxe do
         begin
           if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then LeSet := LeSet + 'G_VENTILABLE' + IntToStr(i) + '=' + '"X" ,'
         end;
       end
       else
       begin
         for i := 1 to MaxAxe do
         begin
           LeSet := LeSet + 'G_VENTILABLE' + IntToStr(i) + '=' + '"-" ,'
         end;
       end;
     end;
     if(Champ='G_DATEMODIF') And (LeSet<>'') then
        BEGIN LeSet:=LeSet+Champ+'='+Valeur ; Continue ; END ;
     if (Champ='G_TPF') or (Champ='G_TVA') or (Champ='G_TVASURENCAISS') then
        BEGIN
        if(Ctrl='CHA') or (Ctrl='PRO') or (Ctrl='IMO')then LeSet:=LeSet+Champ+'='+Valeur ;
        END else
            if(Champ='G_MODEREGLE') or (Champ='G_JOURPAIEMENT1')or
              (Champ='G_JOURPAIEMENT2') or (Champ='G_REGIMETVA')or
              (Champ='G_TVAENCAISSEMENT') or (Champ='G_SOUMISTPF')or
              (Champ='G_RISQUETIERS') or (Champ='G_RELANCEREGLEMENT') or
              (Champ='G_LETTREPAIEMENT') or (Champ='G_RELANCETRAITE') or
              (Champ='G_MOTIFVIREMENT')then
              BEGIN
              if(Ctrl='TIC') or (Ctrl='TID') then LeSet:=LeSet+Champ+'='+Valeur;
              END
            // Test nature pour modif Champ IAS14  // Modif IFRS
            else if ( Champ = 'G_IAS14' ) then
              begin
              if pos( Ctrl , 'COF;COC;COS;COD;DIV;TIC;TID' ) > 0 then
                LeSet := LeSet + Champ + '=' + Valeur ;
              end
            // Fin Modif IFRS
            else
             if (Pos('G_CUTOFF',Champ)>0) or (Pos('G_CUTOFFECHUE',Champ)>0)or (Pos('G_CUTOFFPERIODE',Champ)>0) then
               begin

                if pos( Ctrl , 'CHA;PRO' ) > 0 then
                 LeSet := LeSet + Champ + '=' + Valeur ;

               end // cutoff
                else
                 if (Champ<>'G_DATEMODIF') then
                  LeSet:=LeSet+Champ+'='+Valeur ;

     END ;
  if Table='TIERS' then
  BEGIN
     if(Champ='T_DATEMODIF') And (LeSet<>'') then
        BEGIN LeSet:=LeSet+Champ+'='+Valeur ; Continue ; END ;
     if(Champ='T_TVAENCAISSEMENT')then
       BEGIN
       if (Ctrl='FOU') Or (Ctrl='AUC') then LeSet:=LeSet+Champ+'='+Valeur ;
       END else
     if (Champ='T_RELEVEFACTURE') or (Champ='T_FREQRELEVE') or
        (Champ='T_JOURRELEVE') Then
        BEGIN
        If (Ctrl='CLI') or(Ctrl='AUD') Then LeSet:=LeSet+Champ+'='+Valeur ;
        END Else
     if (Champ='T_RELANCEREGLEMENT') or (Champ='T_RELANCETRAITE') Then
        BEGIN
        If (Ctrl='CLI') or(Ctrl='AUD') Then LeSet:=LeSet+Champ+'='+Valeur ;
        END Else
     if (Champ='T_LETTREPAIEMENT') Then
        BEGIN
        If Ctrl<>'SAL' Then LeSet:=LeSet+Champ+'='+Valeur ;
        END Else
     if (Champ='T_MOTIFVIREMENT') Then
        BEGIN
        If Ctrl<>'SAL' Then LeSet:=LeSet+Champ+'='+Valeur ;
        END Else
     {b FP 13/10/2005 FQ16855: Ce sont des champs montants:
           Retire les guillemets et remplace la virgule par un point}
     if (Champ='T_CREDITPLAFOND') or (Champ='T_CREDITDEMANDE') or (Champ='T_CREDITACCORDE') Then
       BEGIN
       s          := Copy(Valeur, 2, Length(Valeur)); {Retire le 1er guillemet}
       ValeurReel := Copy(s, 1, Pos('"', s)-1);
       Valeur   := Copy(s, Pos('"', s)+1, Length(s));
       ValeurReel := StrFPoint(HEnt1.Valeur(ValeurReel));
       LeSet:=LeSet+Champ+'='+ValeurReel+Valeur ;
       END Else
     {e FP 13/10/2005 FQ16855}
     if(Champ<>'T_DATEMODIF')then LeSet:=LeSet+Champ+'='+Valeur;
     if (Champ = 'T_MULTIDEVISE') then   // ajout me
     begin
          if (Valeur = '"X" ,') then
          LeSet := LeSet + 'T_DEVISE="" ,'
          else
          if (Valeur = '"-" ,') then
          LeSet := LeSet + 'T_DEVISE="'+V_PGI.DevisePivot+'" ,';
     end;
     END ;
  if(Table='SECTION')then LeSet:=LeSet+Champ+'='+Valeur ;
  if(Table='BUDGENE') then LeSet:=LeSet+Champ+'='+Valeur ;
  if(Table='ARTICLE') then LeSet:=LeSet+Champ+'='+Valeur ;
  if(Table='BUDSECT') then LeSet:=LeSet+Champ+'='+Valeur ;
  if(Table='BUDJAL') then LeSet:=LeSet+Champ+'='+Valeur ;

  if(Table = 'ECRITURE') then LeSet := LeSet + Champ + ' = ' + Valeur ;

  if(Table <> 'GENERAUX') And (Table <> 'SECTION') And (Table <> 'TIERS') And
    (Table <> 'BUDGENE')  And (Table <> 'ARTICLE') And (Table <> 'BUDSECT') And
    (Table <> 'BUDJAL') And (Table <> 'ECRITURE') then
    BEGIN
      Result := '' ;
      Exit ;
    END ;
  END;

  if LeSet <> '' then
    Result := LeUpdate + LeSet ;
END;

end.
