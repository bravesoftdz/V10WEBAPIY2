unit TofRubImport;

interface

uses Classes, StdCtrls, SysUtils,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     Dialogs, controls, DB, Forms,
     UTof, UTob, HCtrls, HEnt1, Mul, HTB97, hmsgbox, HStatus,
     FE_Main ;

procedure CPLanceFiche_RubImport;

type
  TOF_RubImport = class(TOF)
  private
    FName: THEdit;
    Tous, Manquante, Existante: TRadioButton;
    TobFams, TobRubs: Tob;
    BHelp: TToolbarButton97;
    procedure BHelpClick(Sender: TObject);
    procedure RubImport(const FileName: string);
    procedure FNameElipsisClick(Sender : TObject);
    procedure FileNameExit(Sender: TObject) ;
    function ImporTob(aTob: Tob): integer;
  public
    procedure OnUpdate ; override ;
    procedure OnLoad ; override ;
  end ;

implementation

{ TOF_RubImport }

procedure CPLanceFiche_RubImport;
begin
  AGLLanceFiche('CP', 'RUBIMPORT', '', '', '');
end;

procedure TOF_RubImport.OnLoad;
begin
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  FName:=THEdit(GetControl('FILENAME'));
  Tous:=TRadioButton(GetControl('TOUS'));
  Manquante:=TRadioButton(GetControl('MANQUANTE'));
  Existante:=TRadioButton(GetControl('EXISTANTE'));
  Ecran.HelpContext:=3221000;
  if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
  if FName<>nil then begin
  if ctxPCL in V_PGI.PGIContexte then FName.Text:=ExtractFilePath(Application.ExeName)+'Rubrique.txt' Else
                                      FName.Text:='C:\PGI00\STD\Rubrique.txt' ;
    FName.OnElipsisClick:=FNameElipsisClick;
    FName.OnExit:=FileNameExit;
    SetFocusControl(FName.Name);
  end;
end;

procedure TOF_RubImport.OnUpdate;
var
  i: integer;
begin
  if Manquante.Checked then
    i := HShowMessage('0;Ajout des rubriques manquantes;Confirmez-vous l''ajout des rubriques manquantes ?;Q;YN;Y;N;;;','','')
  else i := HShowMessage('0;ATTENTION !;Cette mise à jour effacera vos paramétrages antérieurs. Confirmez-vous la mise à jour ?;Q;YN;Y;N;;;','','');
  if i=mrYes then
  begin
    RubImport(FName.Text);
    PGIInfo('     Traitement terminé','Mise à jour des rubriques');
  end;
end;

function TOF_RubImport.ImporTob(aTob: Tob): integer;
begin
  TobRubs:=aTob;
  result:=0;
end;

procedure TOF_RubImport.RubImport(const FileName: string);
var i: integer;
    QFam, QRub: TQuery ;
    where: string;
    Existe: boolean;
BEGIN
  FileNameExit(FName);
  QFam:=nil;
  QRub:=nil;
  where:='WHERE ' + 'RB_RUBRIQUE LIKE "BIA%" OR '
                  + 'RB_RUBRIQUE LIKE "BIP%" OR '
                  + 'RB_RUBRIQUE LIKE "CRC%" OR '
                  + 'RB_RUBRIQUE LIKE "CRP%" OR '
                  + 'RB_RUBRIQUE LIKE "SIG%"';
try
//  QFam:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RBF" AND CC_CODE>="001" AND CC_CODE<="006"',True) ;
  QFam:=OpenSql('SELECT * FROM CHOIXDPSTD WHERE YDS_TYPE="RBF" AND YDS_CODE>="001" AND YDS_CODE<="006"',True) ;
  QRub:=OpenSql('SELECT * FROM RUBRIQUE ' + where + ' ORDER BY RB_RUBRIQUE',True) ;
//  TobRubs:=TOB.Create('RUB_S',nil,-1);
  TOBLoadFromFile(FileName,ImporTob);
  TobFams:=TOB.Create('FAM_S',nil,-1);
  TobFams.Dupliquer(TobRubs,true,true);
  TobFams.Detail[0].ClearDetail;
(*
  if Tous.Checked then begin //Mise à jour de toutes les rubriques
    TobFams.InsertOrUpdateDB(true);
    TobRubs.InsertOrUpdateDB(true);
    Exit;
  end;
*)
  InitMove(TobFams.Detail.Count+TobRubs.Detail[0].Detail.Count,'');
  for i:=0 to TobFams.Detail.Count-1 do begin
    MoveCur(false);
    Existe:=QFam.Locate('YDS_TYPE;YDS_CODE',
                        VarArrayOf(['RBF', TobFams.Detail[i].GetValue('YDS_CODE')]),
                        [loCaseInsensitive]);
    TobFams.Detail[i].PutValue('YDS_PREDEFINI','DOS') ;
    TobFams.Detail[i].PutValue('YDS_NODOSSIER',V_PGI.NoDossier) ;
    if Tous.Checked then //Mise à jour de toutes les rubriques
      TobFams.Detail[i].InsertOrUpdateDB(false);
    if Manquante.Checked and (not Existe) then   //Ajout des rubriques manquantes
      TobFams.Detail[i].InsertDB(nil);
    if Existante.Checked and Existe then //Mise à jour des rubriques existantes
      TobFams.Detail[i].UpdateDB;
  end;
//    InitMove(TobRubs.Detail[0].Detail.Count,'');
  for i:=0 to TobRubs.Detail[0].Detail.Count-1 do begin
    MoveCur(false);
    Existe:=QRub.Locate('RB_RUBRIQUE', TobRubs.Detail[0].Detail[i].GetValue('RB_RUBRIQUE'), [loCaseInsensitive]);
    TobRubs.Detail[0].Detail[i].PutValue('RB_PREDEFINI','DOS') ;
    TobRubs.Detail[0].Detail[i].PutValue('RB_NODOSSIER',V_PGI.NoDossier) ;
    if Tous.Checked then //Mise à jour de toutes les rubriques
      TobRubs.Detail[0].Detail[i].InsertOrUpdateDB(false);
    if Manquante.Checked and (not Existe) then   //Ajout des rubriques manquantes
      TobRubs.Detail[0].Detail[i].InsertDB(nil);
    if Existante.Checked and Existe then //Mise à jour des rubriques existantes
      TobRubs.Detail[0].Detail[i].UpdateDB;
  end;
  FiniMove;
finally
  Ferme(QFam);
  Ferme(QRub);
//  TobRubs.Free;
end;
END;

procedure TOF_RubImport.FNameElipsisClick(Sender: TObject);
begin
  with TOpenDialog.Create(nil) do
  try
//    InitialDir:=ExtractFilePath(Application.ExeName);
    DefaultExt:='TXT';
    Filter:='Fichier Texte (*.txt)|*.txt';
    if Execute then FName.Text:=FileName;
  finally
    Free;
  end;
end;

procedure TOF_RubImport.FileNameExit(Sender: TObject);
var sName: string;
begin
  if trim(GetControlText(TControl(Sender).Name))='' then begin
    HShowMessage('0;;Nom de fichier est obligatoire;E;O;O;O;','','') ;
    SetFocusControl(TControl(Sender).Name);
    Abort;
    Exit;
  end;
  sName:=trim(GetControlText(TControl(Sender).Name));
  if not FileExists(sName) then begin
    HShowMessage('0;;Fichier inexistant;E;O;O;O;','','') ;
    SetFocusControl(TControl(Sender).Name);
    Abort;
  end;
end;


procedure TOF_RubImport.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

initialization
RegisterClasses([TOF_RubImport]);

end.
