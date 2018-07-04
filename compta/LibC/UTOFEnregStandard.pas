unit UTOFEnregStandard;

interface

uses
    Windows,
    Messages,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    classes,
    sysutils,
    UTOF,
    HMsgBox,
    HTB97,
    ParamSoc,
    Grids,
    HCtrls,
    UTob,
    AGLInitPlus,
    AssistPL,
    HEnt1,
    CLgCpte,
    Spin,
    {$IFDEF EAGLCLIENT}
    MainEagl,
    uLanceProcess,
    {$ELSE}
    FE_Main,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$ENDIF}
    uLibStdCpta;

procedure CPLanceFiche_EnregStd;

type
  TOF_ENREGSTANDARD = Class (TOF)
     procedure OnLoad ; override ;
     procedure OnUpdate ; override ;
     procedure OnArgument (S : String ) ; override ;
     private
       NumMax  : integer;
       procedure SupprimeDonneStd (TableStd : string; numstd : string);
       procedure OnClickNum(Sender: TObject);
  end;

const
     TexteMessage : Array[1..1] of String = (
          {1} 'Enregistrement impossible.'
                                          );
implementation

procedure CPLanceFiche_EnregStd;
begin
  AGLLanceFiche('CP', 'ENREGSTANDARD', '', '', ''); 
end;

procedure TOF_ENREGSTANDARD.OnArgument (S : String ) ;
begin
  Inherited ;
  TSpinEdit (GetControl('NUMSTD')).OnExit := OnClickNum;
end;

procedure TOF_ENREGSTANDARD.OnLoad ;
var T, ATob : TOB;
    QPlanRef : TQuery;
    Num      : integer;
begin
  Num := 21;
  QPlanRef := OpenSQL('SELECT * FROM STDCPTA', True);
  {$IFDEF EAGLCLIENT}
  THGrid(GetControl('LISTESTD')).RowCount := QPlanRef.Detail.Count + 1;
  {$ELSE}
  THGrid(GetControl('LISTESTD')).RowCount := QCount(QPlanRef) + 1;
  {$ENDIF}
  T := TOB.Create('', nil, -1);
  while not QPlanRef.Eof do
  begin
    ATob := TOB.Create('STANDARD',T,-1);
    ATob.AddChampSup('No',True);
    ATob.PutValue('No',QPlanRef.FindField('STC_NUMPLAN').AsString);
    ATob.AddChampSup('Libelle',True);
    ATob.PutValue('Libelle',QPlanRef.FindField('STC_LIBELLE').AsString);
// ajout me
    Num := QPlanRef.findField('STC_NUMPLAN').AsInteger;
    QPlanRef.Next;
  end;

  // ajout me
  if Num >= 21 then
  begin
     SetControlText ('NUMSTD', IntToStr (Num+1));
     NumMax := Num+1;
  end else NumMax := 21;

  Ferme (QPlanRef);
  T.PutGridDetail(THGrid(GetControl('LISTESTD')),False,False,'No;Libelle');
  T.Free;
end;

procedure TOF_ENREGSTANDARD.OnUpdate;
var
QStd,QParamSocRef      : Tquery;
Texte                  : string;
lggen, lgaux           : integer;
{$IFDEF EAGLCLIENT}
lTobParam,lTobResult   : TOB;
{$ENDIF}
lstParamStd            : string;
begin
// ajout me
  if GetControlText('LIBELLESTD') = '' then
  begin
    SetFocusControl('LIBELLESTD');
    LastError := 1;
    LastErrorMsg := 'Il faut renseigner un libellé';
    exit;
  end;
  if StrToInt (GetControlText ('NUMSTD')) <= 20 then
  begin
    SetFocusControl('NUMSTD');
    LastError := 1;
    LastErrorMsg := 'Numéro de Standard réservé à CEGID';
    exit;
  end;

  if StrToInt (GetControlText ('NUMSTD')) > 99 then
  begin
    SetFocusControl('NUMSTD');
    LastError := 1;
    LastErrorMsg := 'Numéro de Standard doit être < à 100';
    exit;
  end;

lggen := 0;
lgaux := 0;
  QStd := OpenSQL ('SELECT * FROM STDCPTA WHERE STC_NUMPLAN='+GetControlText('NUMSTD'),False);
  if not QStd.Eof then
  begin
      Texte := 'Le plan existe déjà voulez-vous le remplacer';
      if HShowMessage('0;Chargement du dossier type;'+Texte+';Q;YN;N;N;','','')<>mrYes then
      begin Ferme (QStd); exit ; end;
      if GetControlText('PARAMSTD')='X' then SupprimeDonneStd ('PARSOCREF',GetControlText('NUMSTD'));
      if GetControlText('COMPTESTD')='X' then SupprimeDonneStd ('GENERAUXREF',GetControlText('NUMSTD'));
      if GetControlText('AUXILIAIRESTD')='X' then SupprimeDonneStd ('TIERSREF',GetControlText('NUMSTD'));
      if GetControlText('JOURNALSTD') = 'X' then SupprimeDonneStd ('JALREF',GetControlText('NUMSTD'));
      if GetControlText('GUIDESTD')='X' then
      begin
           SupprimeDonneStd ('GUIDEREF',GetControlText('NUMSTD'));
           SupprimeDonneStd ('ECRGUIREF',GetControlText('NUMSTD'));
           SupprimeDonneStd ('ANAGUIREF',GetControlText('NUMSTD'));
      end;
      QParamSocRef := OpenSQL ('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
                      GetControlText('NUMSTD') + ' AND PRR_SOCNOM="SO_LGCPTEGEN"',True) ;
                      if not QParamSocRef.EOF then
     lggen := QParamSocRef.FindField('PRR_SOCDATA').Asinteger;
     ferme (QParamSocRef);

     QParamSocRef := OpenSQL ('SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN = ' +
                      GetControlText('NUMSTD') + ' AND PRR_SOCNOM="SO_LGCPTEAUX"',True) ;
     if not QParamSocRef.EOF then
     lgaux := QParamSocRef.FindField('PRR_SOCDATA').Asinteger;
     ferme (QParamSocRef);

  end
  else
      if GetControlText('PARAMSTD') <> 'X' then  SetControlText('PARAMSTD', 'X');
  Ferme (QStd);

  { Positionnement des flags de mise à jour des tables }
  if GetControlText('PARAMSTD') = 'X' then lstParamStd := 'X'
  else lstParamStd := '-';
  if GetControlText('COMPTESTD')='X' then lstParamStd := lstParamStd + 'X'
  else lstParamStd := lstParamStd + '-';
  if GetControlText('AUXILIAIRESTD')='X' then lstParamStd := lstParamStd + 'X'
  else lstParamStd := lstParamStd + '-';
  if GetControlText('JOURNALSTD')='X' then lstParamStd := lstParamStd + 'X'
  else lstParamStd := lstParamStd + '-';
  if GetControlText('GUIDESTD')='X' then lstParamStd := lstParamStd + 'X'
  else lstParamStd := lstParamStd + '-';

  {$IFDEF EAGLCLIENT}
  lTobParam := InitTobParamProcessServer;
  lTobParam.AddChampSupValeur('NUMSTD',StrToInt(GetControlText('NUMSTD')));
  lTobParam.AddChampSupValeur('LIBELLESTD',GetControlText('LIBELLESTD'));
  lTobParam.AddChampSupValeur('PARAMSTD',lstParamStd);
  lTobResult := LanceProcessServer('cgiStdCpta', 'SAVEDOSSIERASSTANDARD', GetControlText('NUMSTD'), lTobParam, True ) ;
  if lTobResult.GetValue('RESULT') = 'PASOK' then
  {$ELSE}
  if SaveDossierAsStandard( StrToInt(GetControlText('NUMSTD')), lstParamStd,GetControlText('LIBELLESTD')) <> 0 then
  {$ENDIF}
  begin
    SetFocusControl('NUMSTD');
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end;
  if (lggen <> 0) and (lggen <> Getparamsoc('SO_LGCPTEGEN')) and (GetControlText('PARAMSTD')='X') then
      UpdateChampCompte ( 'ECRGUIREF', ['EGR_GENERAL'],'0', Getparamsoc('SO_LGCPTEGEN'),'GEN',nil);
  if (lgaux <> 0) and (lgaux <> Getparamsoc('SO_LGCPTEAUX')) and (GetControlText('PARAMSTD')='X') then
      UpdateChampCompte ( 'ECRGUIREF', ['EGR_AUXILIAIRE'],'0', Getparamsoc('SO_LGCPTEAUX'),'AUX',nil);
    PGIInfo('Le standard a été correctement enregistré','Enregistrer comme standard');
end;

procedure TOF_ENREGSTANDARD.SupprimeDonneStd (TableStd : string; numstd : string);
var QDos : TQuery;
PrefStd  : String;
Where    : string;
begin
  PrefStd := TableToPrefixe(TableStd);

Where := ' WHERE '+PrefStd+'_NUMPLAN='+numstd;

  QDos := OpenSQL('SELECT * FROM '+TableStd + Where,True);
  if not QDos.Eof then
      ExecuteSQL('Delete from '+TableStd + Where);
  Ferme (QDos);
end;

procedure TOF_ENREGSTANDARD.OnClickNum(Sender: TObject);
var
   Q1    : Tquery;
begin
            Q1 := OpenSql ('SELECT STC_LIBELLE FROM STDCPTA WHERE STC_NUMPLAN='+ GetControlText('NUMSTD'), TRUE);
            if not Q1.EOF then
               SetControlText('LIBELLESTD', Q1.Fields[0].asstring);
            ferme (Q1);
            if StrToInt (GetControlText ('NUMSTD')) <= 20 then
            begin
              PgiInfo ('Numéro de Standard réservé à CEGID', 'STANDARD CEGID');
              SetFocusControl('NUMSTD');
              SetControlText ('NUMSTD', IntToStr(NumMax));
            end;

end;


Initialization
RegisterClasses([TOF_ENREGSTANDARD]);
end.
