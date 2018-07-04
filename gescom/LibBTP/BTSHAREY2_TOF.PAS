{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/09/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSHAREY2 ()
Mots clefs ... : TOF;BTSHAREY2
*****************************************************************}
Unit BTSHAREY2_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97,
     Windows,
     messages,
     UTOF,
     URegenVues ;

Type
  TOF_BTSHAREY2 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure EcritDatas;
    procedure BdeleteClick (Sender : Tobject);
  end ;

Implementation

uses Math, TntStdCtrls,Paramsoc;

procedure TOF_BTSHAREY2.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSHAREY2.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSHAREY2.OnUpdate ;
begin
  Inherited ;
  if V_PGI.DBName = THValComboBox (GetControl('CBLISTDB')).text then
  begin
    PGIError ('Vous ne pouvez pas partager la base avec elle même ...');
    ecran.ModalResult := 0;
    exit;
  end;
  if PGIAsk('ATTENTION : Vous allez partager les données avec une base de donnée CEGID Y2.#13#10 Confirmez-vous ?') <> mryes then
  begin
    ecran.ModalResult := 0;
    exit;
  end;
  if PGIAsk('Confirmez-vous la mise en place avec la base de donnée sélectionnée ?') <> mryes then
  begin
    ecran.ModalResult := 0;
    exit;
  end;
  if TRANSACTIONS (EcritDatas,0) = OeOk then
  begin
//    RegenereVues;
    PgiInfo ('Les données sont à présent partagés avec la base de donnée CEGID Y2 '+THValComboBox (GetControl('CBLISTDB')).Value+'.#13#10 Veuillez regénérer les vues en vous reconnectant à l''applicatif');
    PostMessage(Application.Handle, WM_CLOSE, 0, 0);
  end else
  begin
    PgiInfo ('ERREUR d''inscription du partage');
    ecran.Close;
  end;
end ;

procedure TOF_BTSHAREY2.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSHAREY2.OnArgument (S : String ) ;
var QQ : TQuery;
    Y2Share : string;
    IndexSel,CurIndex : Integer;
    FindItem : boolean;
begin
  Inherited ;
  IndexSel := 1;
  Y2Share := GetParamSocSecur('SO_BTSHAREY2','');
  if ExisteSQL('SELECT 1 FROM DESHARE') then
  begin
    TToolbarButton97 (GetControl('BDelete')).Visible := true;
    TToolbarButton97 (GetControl('BDelete')).OnClick := BdeleteClick;
  end;
  QQ := OpenSql ('SELECT NAME FROM SYSDATABASES WHERE DBID > 4',true,-1,'',true);
  if not QQ.eof then
  begin
    QQ.first;
    CurIndex := 0;
    FindItem := false;
    repeat
      THValComboBox (GetControl('CBLISTDB')).AddItem(QQ.fields[0].AsString,nil);
      if (Y2Share <> '') and (QQ.fields[0].AsString=Y2Share) then
      begin
        IndexSel := CurIndex;
        FindItem := True;
      end;
      qq.next;
      inc(CurIndex);
    until QQ.eof;
    if FindItem then THValComboBox (GetControl('CBLISTDB')).ItemIndex := IndexSel;
  end;
  ferme (QQ);
end ;

procedure TOF_BTSHAREY2.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSHAREY2.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSHAREY2.OnCancel () ;
begin
  Inherited ;                         
end ;

procedure TOF_BTSHAREY2.EcritDatas;
var DbName,SQL : string;
    TOBDatas : TOB;
begin
  DBName := THValComboBox (GetControl('CBLISTDB')).Text;
{$IFDEF DIRECT}
  // PARAMSOC
  ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="'+DBName+'" WHERE SOC_NOM="SO_BTSHAREY2"');
  //
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;001;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;027;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;035;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;012;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;002;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;023;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;006;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;005;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;014;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;031;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;013;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;018;%"') = 0 then V_PGI.IOError := OeUnknown;
  if ExecuteSQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select soc_nom,"LIB","'+DBName+'","PAR","" from paramsoc where soc_tree like "001;021;%"') = 0 then V_PGI.IOError := OeUnknown;
  // TABLES
  SQL := 'insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) '+
  			 'select dt_nomtable ,"LIB","'+DBName+'","TAB","" '+
         'from detables '+
         'where dt_domaine IN ("C","P","Y","T","D","0") AND DT_NOMTABLE NOT IN ("RTDOCUMENT","RESSOURCEEQUIPE","RESSOURCEPR","CHOIXCOD","CHOIXEXT","EXERCICE","YMYBOBS","VENTEANA","VENTIL")';
  if EXECUTESQL(SQL)=0 then V_PGI.IOError := OeUnknown;
  // Tablettes CHOIXCOD et CHOIXEXT
  if EXECUTESQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) select do_combo ,"LIB","'+DBName+'","TTE","" from decombos where do_domaine in ("C","P","Y","T","D","0")  and do_prefixe in ("CC","YX")')=0 then V_PGI.IOError := OeUnknown;
  // EXERCICE
  if EXECUTESQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) values ("EXERCICE","LIB","'+DBName+'","TAB","")')=0 then V_PGI.IOError := OeUnknown;
  // SOCIETE
  if EXECUTESQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) values ("SOCIETE","LIB","'+DBName+'","TAB","")')=0 then V_PGI.IOError := OeUnknown;
  // COMPTEURS
  if EXECUTESQL('insert into deshare (ds_nomtable,ds_modefonc,ds_nombase,ds_typtable,ds_vue) values ("DESEQUENCES","LIB","'+DBName+'","TAB","")')=0 then V_PGI.IOError := OeUnknown;
{$ELSE}
  TOBDatas := TOB.Create('LES DATAS',nil,-1);
  TRY
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;001;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,false);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;027;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;035;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;012;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;023;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;006;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;005;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;014;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;031;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;013;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;018;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select soc_nom as ds_nomtable,"LIB" as ds_modefonc,"'+DBName+'" as ds_nombase,"PAR" as ds_typtable,"" as ds_vue from paramsoc where soc_tree like "001;021;%"';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'SELECT DT_NOMTABLE as ds_nomtable ,"LIB" AS ds_modefonc,"'+DBName+'" as ds_nombase ,"TAB" as ds_typtable,"" as ds_vue from detables '+
           'where dt_domaine IN ("C","P","Y","T","D","0") AND DT_NOMTABLE NOT IN ("RTDOCUMENT","RESSOURCEEQUIPE","RESSOURCEPR","CHOIXCOD","CHOIXEXT","EXERCICE","YMYBOBS")';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'select do_combo as ds_nomtable,"LIB" AS ds_modefonc,"'+DBName+'" as ds_nombase,"TTE" as ds_typtable,"" as ds_vue from decombos where do_domaine in ("C","P","Y","T","D","0")  and do_prefixe in ("CC","YX")';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'SELECT "EXERCICE" as ds_nomtable ,"LIB" AS ds_modefonc,"'+DBName+'" as ds_nombase,"TAB" as ds_typtable,""  as ds_vue';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'SELECT "SOCIETE" as ds_nomtable ,"LIB" AS ds_modefonc,"'+DBName+'" as ds_nombase,"TAB" as ds_typtable,""  as ds_vue';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    SQL := 'SELECT "DESEQUENCES" as ds_nomtable ,"LIB" AS ds_modefonc,"'+DBName+'" as ds_nombase,"TAB" as ds_typtable,""  as ds_vue';
    TOBDatas.LoadDetailDBFromSQL ('DESHARE',SQL,true);
    TOBDatas.SetAllModifie(true);
    TOBDatas.InsertOrUpdateDB(false);
    // PARAMSOC
    ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="'+DBName+'" WHERE SOC_NOM="SO_BTSHAREY2"');
    //
  FINALLY
    TOBDatas.Free;
  END;
{$ENDIF}
//  RegenereVues;
end;

procedure TOF_BTSHAREY2.BdeleteClick(Sender: Tobject);
begin
  if PGIAsk('ATTENTION : Vous allez casser le lien entre la base de donnée Y2 et cette base de donnée.#13#10 Confirmez-vous cette action ?')<>mryes then exit;
  if PGIAsk('Confirmez-vous la suppression de ce lien ?')<>mryes then exit;
  if GetParamSocSecur('SO_BTSHAREY2','')<> '' then
  begin
    ExecuteSQL('DELETE FROM DESHARE WHERE DS_NOMBASE="'+GetParamSocSecur('SO_BTSHAREY2','')+'"');
    PgiInfo ('Les données ne sont plus partagés avec la base de donnée CEGID Y2 '+THValComboBox (GetControl('CBLISTDB')).Value+'.#13#10 Veuillez vous deconnecter de l''applicatif');
    ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="" WHERE SOC_NOM="SO_BTSHAREY2"');
    PostMessage(Application.Handle, WM_CLOSE, 0, 0);
  end else
  begin
    PgiInfo('Vous ne pouvez réaliser cette opération de cette façon.Merci de procéder manuellement');
  end;
end;

Initialization
  registerclasses ( [ TOF_BTSHAREY2 ] ) ;
end.
