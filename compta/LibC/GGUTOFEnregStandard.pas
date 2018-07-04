unit UTOFEnregStandard;

interface

uses
    Windows, Messages, Graphics, Controls, Forms, Dialogs,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},classes, sysutils,UTOF, HMsgBox, HTB97, ParamSoc, Grids, HCtrls,
    UTob, AGLInitPlus,AssistPL, HEnt1,CLgCpte, Spin;

type
  TOF_ENREGSTANDARD = Class (TOF)
     procedure OnLoad ; override ;
     procedure OnUpdate ; override ;
     procedure OnArgument (S : String ) ; override ;
     private
       NumMax  : integer;
       procedure SaveAsStandard;
       procedure SaveInfoStandard(NumStd : integer; LibelleStd : string);
       procedure SupprimeDonneStd (TableStd : string; numstd : string);
       procedure OnClickNum(Sender: TObject);
  end;

const
     TexteMessage : Array[1..1] of String = (
          {1} 'Enregistrement impossible.'
                                          );
implementation

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
  THGrid(GetControl('LISTESTD')).RowCount := QCount(QPlanRef) + 1;
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


  if Transactions (SaveAsStandard,3) <> oeOK then
  begin
    SetFocusControl('NUMSTD');
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end else
  begin
// ajout me
if (lggen <> 0) and (lggen <> Getparamsoc('SO_LGCPTEGEN')) and (GetControlText('PARAMSTD')='X') then
      MiseajourCpte ('ECRGUIREF', 'EGR_GENERAL' , '0', Getparamsoc('SO_LGCPTEGEN'), 1,FALSE);
if (lgaux <> 0) and (lgaux <> Getparamsoc('SO_LGCPTEAUX')) and (GetControlText('PARAMSTD')='X') then
      MiseajourCpte ('ECRGUIREF', 'EGR_AUXILIAIRE' , '0', Getparamsoc('SO_LGCPTEAUX'), 1,FALSE);
    PGIInfo('Le standard a été correctement enregistré','Enregistrer comme standard');
  end;

end;

procedure TOF_ENREGSTANDARD.SaveAsStandard;
var NumStd : integer;
begin
  NumStd := StrToInt(GetControlText('NUMSTD'));
  // Enregistrement des paramètres dossier
  if GetControlText('PARAMSTD')='X' then SaveParamAsStandardCompta (NumStd);
  // Enregistrement du plan comptable
  if GetControlText('COMPTESTD')='X' then
  begin
       SaveAsStandardCompta(NumStd, 'GENERAUX','GENERAUXREF');
{ Ajout ME le 11/07/2001 }
       ExecuteSQl ('Update GENERAUXREF set GER_TOTDEBP=0,GER_TOTCREP=0,GER_TOTDEBE=0,GER_TOTCREE=0,'+
       'GER_TOTDEBS=0,GER_TOTCRES=0 Where GER_NUMPLAN='+IntToStr(NumStd));
  end;
  // Enregistrement des tiers
  if GetControlText('AUXILIAIRESTD')='X' then SaveAsStandardCompta(NumStd, 'TIERS','TIERSREF');
  // Enregistrement des journaux
  if GetControlText('JOURNALSTD')='X' then SaveAsStandardCompta(NumStd, 'JOURNAL','JALREF');
  // Enregistrement des guides
  if GetControlText('GUIDESTD')='X' then
  begin
    SaveAsStandardCompta(NumStd, 'GUIDE','GUIDEREF');
    SaveAsStandardCompta(NumStd, 'ECRGUI','ECRGUIREF');
    SaveAsStandardCompta(NumStd, 'ANAGUI','ANAGUIREF');
  end;
  SaveInfoStandard (NumStd,GetControlText('LIBELLESTD'));

end;

procedure TOF_ENREGSTANDARD.SaveInfoStandard (NumStd : integer; LibelleStd : string);
var  QPlanRef : TQuery;
begin
  QPlanRef := OpenSQL('SELECT * FROM STDCPTA WHERE STC_NUMPLAN='+IntToStr(NumStd), False);
  if QPlanRef.Eof then
  begin
    QPlanRef.Insert;
    InitNew(QPlanRef);
    QPlanRef.FindField('STC_NUMPLAN').AsInteger := NumStd;
    QPlanRef.FindField('STC_LIBELLE').AsString := LibelleStd;
    QPlanRef.FindField('STC_ABREGE').AsString := Copy(LibelleStd,1,17);
    QPlanRef.FindField('STC_PREDEFINI').AsString := 'STD';
    QPlanRef.Post;
  end;
  Ferme (QPlanRef);
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
