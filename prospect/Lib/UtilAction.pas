unit UtilAction;

interface
uses  Utob
{$IFDEF EAGLCLIENT}
      ,MaineAGL
{$ELSE}
      ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main
{$ENDIF}
      , FactUtil, EntGC, Sysutils, HCtrls, HEnt1, HMsgBox,Controls
      ,M3FP,YRESSOURCE,YPlanning,aglinit,UtofRess_Occup,EntRT,ParamSoc,uEntCommun
      ;

Procedure RTAffichePieceLiee( NumChainage : integer; stProduitpgi : String) ;
procedure RTMontreLaPiece( stNat,stSouche : String; iNum,iIndice : integer) ;
procedure RTCreationLiensActions (TobAction : Tob);
function ControleIsFreeYPL (TobControleYplanning : Tob;DateDebut,DateFin : TDateTime;AffOccup : Boolean) : Boolean;
procedure RTMAJYPlanning (TobRessources : Tob;iDatedebut,iDateFin : TDateTime;Libelle,Etat : string);
function RTActEstPlanifiable (TypeAction : string) : Boolean;
{ Creation d'action via une tob externe }
function CreateActionFromTob(Tiers : String; Numero : integer; TheTob: Tob): String;
function CreateOrUpdateActionFromTob(Tiers : String; Numero : integer; TheTob: Tob): String;
function CreateOrUpdateAction(NewAction : boolean; Tiers : String; Numero : integer; TheTob: Tob): String;
function DeleteAction( Tiers : String; Numero : integer; TheTob: Tob): String;
function ConfirmeSurBooking : Integer;
function RTCreateRAC(Const Argument, Separateur: String ) : Variant;
function RTUpdateRAC(Const Argument, Separateur: String) : Variant;
function RTDeleteRAC(Const Argument, Separateur: String ) : Variant;
function InitTobByArgument(const TobData: Tob; Argument: String; const Separateur: String): Integer;
function PCSDureeCombo( Duree : Double ) : string;
procedure AlimTobAction (TobAction,TobActGen:Tob;StOperation: string;noactgen:integer);

implementation

uses  UTom,
      UTomAction,
      {$IFDEF PGISIDE}
      UtilGC,
      {$ENDIF PGISIDE}
      UtilPGI,wCommuns,
      facture;

const
  // libellés des messages
  TexteMessage: array[1..1] of string = (
          {1}       'Des intervenants ne sont pas disponibles sur cette période. #10#13 Confirmez-vous la planification ?'
    );

procedure RTAffichePieceLiee( NumChainage : integer; stProduitpgi : String) ;
var Q : TQuery;
    TobPiece : tob;
begin
    TobPiece := TOB.Create ('les liens', Nil, -1);
    Q:=OpenSQL('SELECT RLC_INDICEG,RLC_NATUREPIECEG,RLC_NUMERO,RLC_PRODUITPGI,RLC_SOUCHE FROM CHAINAGEPIECES WHERE RLC_NUMCHAINAGE='+intToStr(NumChainage),true,-1,'',true);
    TobPiece.LoadDetailDB('CHAINAGEPIECES', '', '',Q ,false);
    Ferme(Q);

    if TobPiece.detail.count = 0 then
      begin
      PGIBox(TraduireMemoire('Aucune pièce associée au chaînage'),TraduireMemoire('Chaînage'));
      end
    else
    if TobPiece.detail.count = 1 then
      begin
      RTMontreLaPiece(TobPiece.detail[0].GetString('RLC_NATUREPIECEG'),TobPiece.detail[0].GetString('RLC_SOUCHE'),
         TobPiece.detail[0].GetInteger('RLC_NUMERO'),TobPiece.detail[0].GetInteger('RLC_INDICEG'));
      end
    else
      AGLLanceFiche ('RT', 'RTCHAINAGEPIECES','RLC_NUMCHAINAGE='+intToStr(NumChainage)+';RLC_PRODUITPGI='+stProduitPgi, '', '');
    TobPiece.free;
end;

procedure RTMontreLaPiece( stNat,stSouche : String; iNum,iIndice : integer) ;
var CleDoc : R_CleDoc;
    Q : TQuery;
    Requete : String;
    {$IFNDEF BTP}
    Param: R_SaisiePieceParam;
    {$ENDIF}
begin
  Requete:='SELECT GP_DATEPIECE from PIECE where GP_NUMERO='+IntToStr(Inum)+' and GP_SOUCHE="'+stSouche+
  '" and GP_INDICEG='+IntToStr(iIndice)+' and GP_NATUREPIECEG="'+stNat+'"';
  Q:=OpenSQL(Requete,true,-1,'',true);
  if Q.EOF then
    begin
    PGIBox(TraduireMemoire('Cette pièce est inexistante : ')+RechDom('GCNATUREPIECEG',stNat , False)+TraduireMemoire(' n° ')+IntToStr(Inum),TraduireMemoire('Chaînage'));
    Ferme(Q);
    end
  else
    begin
    DecodeRefPiece(stNat + ';' + DateTimeToStr(Q.FindField('GP_DATEPIECE').AsDateTime) + ';' + stSouche
     + ';' + IntToStr(Inum) + ';' + IntToStr(iIndice), CleDoc);
    Ferme(Q);
{$IFDEF BTP}
    SaisiePiece(CleDoc, taConsult);
{$ELSE}
    InitRecordSaisiePieceParam(Param);
    SaisiePiece(CleDoc, taConsult, Param);
{$ENDIF}
    end;
end;

procedure RTCreationLiensActions (TobAction : Tob);
var TobRessource,TobCreatYplanning,TYPL : Tob;
    StGuidYRS : string;
    DureeAct : double;
    Heures,Minutes : integer;
    Resultat : integer;
    Heureact : TDateTime;
begin
  // Création de la table ACTIONINTERVENANT
  if TobAction.GetString ('RAC_INTERVENANT') <> '' then
  begin
    TobRessource:=TOB.create ('ACTIONINTERVENANT',NIL,-1);
    TobRessource.SetString('RAI_AUXILIAIRE', TobAction.GetString ('RAC_AUXILIAIRE'));
    TobRessource.SetInteger('RAI_NUMACTION', TobAction.GetInteger ('RAC_NUMACTION'));
    TobRessource.SetString('RAI_RESSOURCE', TobAction.GetString ('RAC_INTERVENANT'));
    TobRessource.SetString('RAI_GUID', AglGetGuid());
    TobRessource.InsertDB(Nil);
    if RTActEstPlanifiable (TobAction.GetString('RAC_TYPEACTION')) then
      begin
      StGuidYRS := GetYRS_GUID('', TobAction.GetString ('RAC_INTERVENANT'), '');
      if (StGuidYRS <> '') then
        begin
        TobCreatYplanning := Tob.create('CREATIONYPLANNING',Nil,-1) ;
        TYPL:=Tob.create('#YPLANNING',TobCreatYplanning,-1);
        TYPL.AddChampSupValeur('YPL_PREFIXE', 'RAI');
        TYPL.AddChampSupValeur('YPL_GUIDYRS', StGuidYRS);
        TYPL.AddChampSupValeur('YPL_GUIDORI', TobRessource.GetString('RAI_GUID'));
        if TobAction.GetDateTime('RAC_HEUREACTION') <> IDate1900 then
          Heureact := TobAction.GetDateTime('RAC_HEUREACTION')
        else
          Heureact := 0;
        TYPL.AddChampSupValeur('YPL_DATEDEBUT', TobAction.GetDateTime('RAC_DATEACTION') + Heureact);
        Dureeact := TobAction.GetDouble('RAC_DUREEACTION');
        Heures:=trunc(Dureeact/60);
        Minutes:=trunc(Dureeact-(Heures*60));
        TYPL.AddChampSupValeur('YPL_DATEFIN', TobAction.GetDateTime('RAC_DATEACTION') + Heureact + EncodeTime(Heures,Minutes,0,0));
        TYPL.AddChampSupValeur('YPL_LIBELLE', TobAction.GetString('RAC_LIBELLE'));
        TYPL.AddChampSupValeur('YPL_ABREGE', TobAction.GetString('RAC_LIBELLE'));
        TYPL.AddChampSupValeur('YPL_STATUTYPL', TobAction.GetString('RAC_ETATACTION'));
        TYPL.AddChampSupValeur('YPL_PRIVE', '-');
        if TobCreatYplanning.Detail.Count <> 0 then
          begin
            Resultat := CreateYPL(TobCreatYplanning);
            if resultat <> 0 then
               PGIBox('Erreur lors de l''enregistrement dans la table YPLANNING. ' + GetLastErrorMsgYPL(resultat), '');
          end;
        TobCreatYplanning.free;
        end;
      end;
    TobRessource.free;
  end;
end;

function ControleIsFreeYPL (TobControleYplanning : Tob;DateDebut,DateFin : TDateTime;AffOccup : Boolean) : Boolean;
var i : integer;
    stWhere,StGuidYRS,stPrefixe : string;
    EstLibre : Boolean;
    TobOccup : Tob;
begin
Result := True;
EstLibre := True;
for i:=0 to TobControleYplanning.detail.count-1 do
  begin
  StGuidYRS := GetYRS_GUID('', TobControleYplanning.detail[i].GetString('RESSOURCE'), '');
  StWhere := 'YPL_GUIDYRS = "' + StGuidYRS +'" AND YPL_GUIDORI <> "' + TobControleYplanning.detail[i].GetString('GUID') + '"';
  if not IsFreeYPL(cStTableGRC, DateDebut,DateFin,stWhere) then
    begin
    TobControleYplanning.detail[i].SetBoolean ('LIBRE',False);
    StWhere := StWhere + ' AND YPL_DATEDEBUT<"' + UsDateTime(DateFin) + '"'
       + ' AND YPL_DATEFIN>"' + UsDateTime(DateDebut) + '"';
    TobOccup:=Tob.create('les Occupations',Nil,-1) ;
    LoadTobYPL (stWhere,TobOccup);
    if TobOccup.Detail.count > 0 then
      begin
      stPrefixe := TobOccup.detail[0].GetString('YPL_PREFIXE');
      if stPrefixe = 'APL' then TobControleYplanning.detail[i].SetString ('MOTIF','Intervention')
      else if stPrefixe = 'PCN' then TobControleYplanning.detail[i].SetString ('MOTIF','Absence')
      else if stPrefixe = 'RAI' then TobControleYplanning.detail[i].SetString ('MOTIF','Action');
      end;
    FreeAndNil (TobOccup);
    EstLibre := False;
    end;
  end;
if EstLibre = True then Exit;
if AffOccup then
  begin
  TheTob := TobControleYplanning;
  RTLanceFiche_Ress_Occup('RT','RTRESS_OCCUP','','','') ;
  TheTob := Nil;
  end;
Result := False;
end;

function RTActEstPlanifiable (TypeAction : string) : Boolean;
var TobTypAct : Tob;
begin
Result := False;
VH_RT.TobTypesAction.Load;
TobTypAct:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TypeAction,'---',0],TRUE) ;
if (TobTypAct <> Nil) and (TobTypAct.GetBoolean('RPA_PLANIFIABLE') = True) then Result := True;
end;

procedure RTMAJYPlanning (TobRessources : Tob;iDatedebut,iDateFin : TDateTime;Libelle,Etat : string);
var TYPL : Tob;
    StGuidYRS : string;
    i,Resultat : integer;
    TobModifYplanning : Tob;
begin
TobModifYplanning := Tob.create('MODIFYPLANNING',Nil,-1) ;
for i:=0 to TobRessources.detail.count-1 do
  begin
  StGuidYRS := GetYRS_GUID('', TobRessources.Detail[i].GetString('RAI_RESSOURCE'), '');
  if (StGuidYRS <> '') then
    begin
    TYPL:=Tob.create('#YPLANNING',TobModifYplanning,-1);
    TYPL.AddChampSupValeur('YPL_PREFIXE', 'RAI');
    TYPL.AddChampSupValeur('YPL_GUIDYRS', StGuidYRS);
    TYPL.AddChampSupValeur('YPL_GUIDORI', TobRessources.Detail[i].GetString('RAI_GUID'));
    if iDateDebut <> IDate1900 then TYPL.AddChampSupValeur('YPL_DATEDEBUT', iDateDebut);
    if iDateFin <> IDate1900 then TYPL.AddChampSupValeur('YPL_DATEFIN', iDateFin);
    if Libelle <> '' then
      begin
      TYPL.AddChampSupValeur('YPL_LIBELLE', Libelle);
      TYPL.AddChampSupValeur('YPL_ABREGE', Libelle);
      end;
    if Etat <> '' then TYPL.AddChampSupValeur('YPL_STATUTYPL', Etat);
    end;
  end;
if TobModifYplanning.Detail.Count <> 0 then
  begin
    Resultat := UpdateYPL(TobModifYplanning);
    if resultat <> 0 then
       PGIBox('Erreur lors de l''enregistrement dans la table YPLANNING. ' + GetLastErrorMsgYPL(resultat), '');
  end;
FreeAndNil (TobModifYplanning);
end;

function ConfirmeSurBooking : Integer;
begin
Result := -2;
if GetParamSocSecur ('SO_YPLSURBOOKING',True) then
  begin
  Result := -1;
  if PGIAsk(TexteMessage[1],'') = mrYes then Result := 0;
  end;
end;

procedure AGLRTMontreLaPiece(parms:array of variant; nb: integer ) ;
begin
RTMontreLaPiece(String(Parms[0]), String(Parms[1]), Integer(Parms[2]), Integer(Parms[3])) ;
end;

//==============================================================================
//  fonction de creation d'un Action à partir d'une tob.
//  elle applique les regles de validation definies dans la tom Action
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateActionFromTob(Tiers : String; Numero : integer; TheTob: Tob): String;
begin
  Result := '';
  { Contrôle existence du tiers }
  if not ExisteSQL('Select T_TIERS from TIERS where T_TIERS="' + Tiers + '"') then
  begin
    Result := TraduireMemoire('Le tiers ') + Tiers + TraduireMemoire(' n''existe pas dans la table : ') + 'TIERS';
    TheTob.AddChampSupValeur('Error', Result);
    Exit;
  end;
  { Contrôle existence de l'action si Numero différent de 0}
  if Numero <> 0 then
    if ExisteSQL('Select RAC_TIERS from ACTIONS where RAC_TIERS="' + Tiers + '" and RAC_NUMACTION=' + IntToStr(Numero)) then
    begin
      Result := CreateOrUpdateActionFromTob(Tiers, Numero, TheTob);
//      Result := 'Modification d''action actuellement impossible';
      TheTob.AddChampSupValeur('Error', Result);
      Exit;
    end;
  Result := CreateOrUpdateAction(True, Tiers, Numero, TheTob);
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

//==============================================================================
//  fonction de modification d'un Action à partir d'une tob.
//  elle applique les regles de validation definies dans la tom Action
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateOrUpdateActionFromTob(Tiers : String; Numero : integer; TheTob: Tob): String;
begin
  Result := '';
  Result := CreateOrUpdateAction(False, Tiers, Numero, TheTob);
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

function CreateOrUpdateAction(NewAction : boolean; Tiers : String; Numero : integer; TheTob: Tob): String;
var
  iChamp: Integer;
  TypeAction, FieldName: String;
  Exist, IsValid: Boolean;
  TobAction, TobLoc: Tob;
  TomAction: Tom;
  Q : TQuery;

  procedure VerifieModifAutorisee;
  var stChInterdits : string;
  begin
    stChInterdits := 'RAC_LIBELLE;RAC_INTERVENANT';
    iChamp := 999;
    while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
    begin
      Inc(iChamp);
      { Vérifie si le champ fait partie des champs non modifiables }
      if Pos(TheTob.GetNomChamp(iChamp), stChInterdits) > 0 then
        if TheTob.GetValue(TheTob.GetNomChamp(iChamp)) <> TobAction.GetValue(TheTob.GetNomChamp(iChamp)) then
        begin
          if Result = '' then
            Result := TraduireMemoire('Les champs suivants sont présents mais non modifiables : ') + TheTob.GetNomChamp(iChamp)
          else
            Result := Result + ', ' + TheTob.GetNomChamp(iChamp);
        end;
    end;
    if Result <> '' then
      TobAction.Free;
  end;

begin
  Result := '';
  if TheTob <> nil then
  begin
    { Crée une TobAction et contrôle l'existence des champs de la tob passée en paramêtre }
    TobAction := Tob.Create('ACTIONS', nil, -1);
    if not NewAction then
    begin
      Q := OpenSQL('Select * from ACTIONS where RAC_TIERS="' + Tiers + '" and RAC_NUMACTION=' + IntToStr(Numero), True,-1,'',true);
      if not Q.EOF then
      begin
        TobLoc := TOB.Create('', nil, -1);
        TobLoc.LoadDetailDB('ACTIONS', '', '', Q, False);
        if TobLoc.Detail.Count > 0 then
          TobAction.Dupliquer(TobLoc.Detail[0], True, True);
        TobLoc.Free;
        VerifieModifAutorisee;
      end
      else
      begin
        NewAction := True;
        Ferme(Q);
      end;
    end;
    if Result = '' then
    begin
      try
        { Recopie la tob à mettre à jour dans la tob Action }
        iChamp := 999; Exist := True;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
        begin
          Inc(iChamp);
          { Vérifie si le champ fait partie de la table Action }
          Exist := TobAction.FieldExists(TheTob.GetNomChamp(iChamp));
        end;
        if Exist then
        begin
          { Vérifie les données en passant par une TomAction }
          if not NewAction then
            TomAction := TOM. Create(Q, nil, False, 'ACTIONS')
          else
            TomAction := CreateTOM('ACTIONS', nil, false, true);
          TomAction.Argument('ORIGINE=PGISIDE;TYPE='+ TypeAction + ';RAC_TIERS=' + Tiers);

          if not NewAction then
            TomAction.LoadRecord
          else
            TomAction.InitTOB(TobAction);
          { Recopie la tob à mettre à jour dans la tob Action }
          iChamp := 999;
          while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
          begin
            Inc(iChamp);
            FieldName := TheTob.GetNomChamp(iChamp);
            TobAction.PutValue(FieldName, TheTob.GetValue(FieldName));
          end;
          try
            IsValid := TomAction.VerifTOB(TobAction);
            Result := TomAction.LastErrorMsg;
          finally
            TomAction.Free;
          end;
          if IsValid then
          begin
            if NewAction then
            begin
              try
                TobAction.InsertDB(nil, False); { Enregistre la TobAction }
                TheTob.Dupliquer(TobAction, False, True);
              except
                on E: Exception do
                begin
                  {$IFDEF PGISIDE}
                  MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table Action'), E.Message);
                  {$ENDIF PGISIDE}
                  Result := E.Message;
                end;
              end;
            end
            else
            begin
              TobAction.InsertOrUpdateDB; { Enregistre la TobAction }
              TheTob.Dupliquer(TobAction, False, True);
            end;
          end;
          if not NewAction then
            Ferme(Q);
        end
        else
          Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'Action';
      finally
        TobAction.Free;
      end;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

function DeleteAction(Tiers : String; Numero : integer; TheTob: Tob): String;
var
  TypeAction: String;
  IsValid: Boolean;
  TobAction, TobLoc: Tob;
  TomAction: Tom;
  Q : TQuery;

begin
  Result := '';
  if TheTob <> nil then
  begin
    { Crée une TobAction et contrôle l'existence des champs de la tob passée en paramêtre }
    TobAction := Tob.Create('ACTIONS', nil, -1);
    Q := OpenSQL('Select * from ACTIONS where RAC_TIERS="' + Tiers + '" and RAC_NUMACTION=' + IntToStr(Numero), True,-1,'',true);
    if not Q.EOF then
    begin
      TobLoc := TOB.Create('', nil, -1);
      TobLoc.LoadDetailDB('ACTIONS', '', '', Q, False);
      if TobLoc.Detail.Count > 0 then
        TobAction.Dupliquer(TobLoc.Detail[0], True, True);
      TobLoc.Free;
    end
    else
    begin
      result := 'Action inexistante';
      Ferme(Q);
    end;

    if Result = '' then
    begin
      try
          TomAction := TOM. Create(Q, nil, False, 'ACTIONS');
          TomAction.Argument('ORIGINE=PGISIDE;TYPE='+ TypeAction + ';RAC_TIERS=' + Tiers);
          TomAction.InitTOB(TobAction);
          TomAction.LoadRecord;
          try
            IsValid := TomAction.DeleteTOB(TobAction);
            Result := TomAction.LastErrorMsg;
          finally
            TomAction.Free;
          end;
          if IsValid then
          begin
              TobAction.DeleteDB;
              TheTob.Dupliquer(TobAction, False, True);
          end;
          Ferme(Q);
      finally
        TobAction.Free;
      end;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

function RTCreateRAC(Const Argument, Separateur: String ) : Variant;
var TobData :tob;
begin
{minimum RAC_TIERS=x;RAC_TYPEACTION=x;RAC_LIBELLE=x;RAC_INTERVENANT=x}
  TobData := Tob.Create(TableToPrefixe('ACTIONS'), nil, -1);
  try
    Result := InitTobByArgument(TobData, Argument, Separateur);
    if (Result = 0) then
      Result:=CreateOrUpdateAction(True,GetArgumentString(Argument, 'RAC_TIERS', True, Separateur),0,TobData);
  finally
    TobData.Free;
  end;
end;

function RTUpdateRAC(Const Argument, Separateur: String) : Variant;
var TobData :tob;
begin
  TobData := Tob.Create(TableToPrefixe('ACTIONS'), nil, -1);
  try
    Result := InitTobByArgument(TobData, Argument, Separateur);
    if not (TobData.FieldExists('RAC_NUMACTION')) or (TobData.GetInteger ('RAC_NUMACTION') = 0) then
       result:=TraduireMemoire('le numéro d''action est obligatoire')
    else
      if (Result = 0) then
        Result:=CreateOrUpdateAction(false,GetArgumentString(Argument, 'RAC_TIERS', True, Separateur),TobData.GetInteger ('RAC_NUMACTION'),TobData);
  finally
    TobData.Free;
  end;
end;

function RTDeleteRAC(Const Argument, Separateur: String ) : Variant;
var TobData :tob;
begin
  TobData := Tob.Create(TableToPrefixe('ACTIONS'), nil, -1);
  try
    Result := InitTobByArgument(TobData, Argument, Separateur);
    if not (TobData.FieldExists('RAC_NUMACTION')) or (TobData.GetInteger ('RAC_NUMACTION') = 0) then
       result:=TraduireMemoire('le numéro d''action est obligatoire')
    else
      if (Result = 0) then
        Result:=DeleteAction(GetArgumentString(Argument, 'RAC_TIERS', True, Separateur),TobData.GetInteger ('RAC_NUMACTION'),TobData);
  finally
    TobData.Free;
  end;

end;

function InitTobByArgument(const TobData: Tob; Argument: String; const Separateur: String): Integer;
var
  Token, FieldName: String;
begin
  Result := 0;
  if Argument <> '' then
  begin
    while (Argument <> '') and (Result = 0) do
    begin
      Token := Trim(ReadTokenPipe(Argument, Separateur));
      FieldName := GetTokenFieldName(Token);
      if ChampToNum(FieldName) > 0 then
      begin
        TobData.AddChampSup(FieldName, False);
        case wGetSimpleTypeField(FieldName) of
          'N': TobData.SetDouble  (FieldName, GetArgumentDouble  (Token, FieldName, Separateur));
          'I': TobData.SetInteger (FieldName, GetArgumentInteger (Token, FieldName, Separateur));
          'D': TobData.SetDateTime(FieldName, GetArgumentDateTime(Token, FieldName, Separateur));
          'B': TobData.SetBoolean (FieldName, GetArgumentBoolean (Token, FieldName, Separateur));
        else
          TobData.SetString(FieldName, GetArgumentString(Token, FieldName, False, Separateur))
        end
      end
    end;
  end
end;

function PCSDureeCombo( Duree : Double ) : string;
var w_duree : double;
begin
w_duree := trunc (Duree);
Result:=FloatTostr(w_duree);
if  length(result)<2 then Result:='00'+Result
else if  length(result)<3 then Result:='0'+Result;
end;

procedure AlimTobAction (TobAction,TobActGen:Tob;StOperation: string;noactgen:integer);
begin
  if Stoperation <> '' then TobAction.PutValue ('RAC_OPERATION',Stoperation)
  else TobAction.PutValue ('RAC_OPERATION',TobActGen.GetValue('RAG_OPERATION'));
  TobAction.PutValue ('RAC_NUMACTGEN',noactgen);
  TobAction.PutValue ('RAC_LIBELLE',TobActGen.GetValue('RAG_LIBELLE'));
  TobAction.PutValue ('RAC_TYPEACTION',TobActGen.GetValue('RAG_TYPEACTION'));
  TobAction.PutValue ('RAC_PRODUITPGI',TobActGen.GetValue('RAG_PRODUITPGI'));
  TobAction.PutValue ('RAC_INTERVENANT',TobActGen.GetValue('RAG_INTERVENANT'));
  //TobAction.PutValue ('RAC_INTERVINT',TobActGen.GetValue('RAG_INTERVENANT')+';');
  TobAction.PutValue ('RAC_DATEACTION',TobActGen.GetValue('RAG_DATEACTION'));
  TobAction.PutValue ('RAC_DATEECHEANCE',TobActGen.GetValue('RAG_DATEECHEANCE'));
  if TobActGen.GetValue('RAG_ETATACTION') <> '' then
    TobAction.PutValue ('RAC_ETATACTION',TobActGen.GetValue('RAG_ETATACTION'));
  TobAction.PutValue ('RAC_TABLELIBRE1',TobActGen.GetValue('RAG_TABLELIBRE1'));
  TobAction.PutValue ('RAC_TABLELIBRE2',TobActGen.GetValue('RAG_TABLELIBRE2'));
  TobAction.PutValue ('RAC_TABLELIBRE3',TobActGen.GetValue('RAG_TABLELIBRE3'));
  TobAction.PutValue ('RAC_TABLELIBREF1',TobActGen.GetValue('RAG_TABLELIBREF1'));
  TobAction.PutValue ('RAC_TABLELIBREF2',TobActGen.GetValue('RAG_TABLELIBREF2'));
  TobAction.PutValue ('RAC_TABLELIBREF3',TobActGen.GetValue('RAG_TABLELIBREF3'));

  TobAction.PutValue ('RAC_COUTACTION',TobActGen.GetValue('RAG_COUTACTION'));
  if TobActGen.FieldExists('RAG_DUREEACTION') then
    TobAction.PutValue ('RAC_DUREEACTION',TobActGen.GetValue('RAG_DUREEACTION'));
  TobAction.PutValue ('RAC_BLOCNOTE',TobActGen.GetValue('RAG_BLOCNOTE'));
end;

Function AGLRTCreateRAC( parms: array of variant; nb: integer ) : variant ;
begin
  result := RTCreateRAC(parms[0],parms[1]);
end;

Function AGLRTUpdateRAC( parms: array of variant; nb: integer ) : variant ;
begin
  result := RTUpdateRAC(parms[0],parms[1]);
end;

Function AGLRTDeleteRAC( parms: array of variant; nb: integer ) : variant ;
begin
  result := RTDeleteRAC(parms[0],parms[1]);
end;

function PCSFormatDuree( parms: array of variant; nb: integer ) : Variant;
var stDuree : string;
begin
stDuree:=string(parms[0]);
result:=Valeur(stDuree);
end;

Initialization
RegisterAglProc( 'RTMontreLaPiece', false , 4, AGLRTMontreLaPiece);
RegisterAglFunc( 'G_CreateRAC', FALSE, 2, AGLRTCreateRAC);
RegisterAglFunc( 'G_UpdateRAC', FALSE, 2, AGLRTUpdateRAC);
RegisterAglFunc( 'G_DeleteRAC', FALSE, 2, AGLRTDeleteRAC);
RegisterAglFunc( 'PCSFormatDuree', False , 1, PCSFormatDuree);
end.
