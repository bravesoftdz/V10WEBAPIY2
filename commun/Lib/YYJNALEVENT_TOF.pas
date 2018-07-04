{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYJNALEVENT ()
Mots clefs ... : TOF;YYJNALEVENT
*****************************************************************}
Unit YYJNALEVENT_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , FE_Main
  {$ENDIF EAGLCLIENT}
  , uTob
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  , CBPMcd
  , Htb97
  , uTOFComm
  , HDB
  ;


Type
  TOF_YYJNALEVENT = Class (tTOFComm)
  private
    lFListe  : THDBGrid;
    VoirDoc : TToolbarButton97;

    procedure FListe_OnRowEnter(Sender: TObject);
    procedure VoirDoc_OnClick(Sender : TObject);
    function GetBFilesKey :  string;
    function GetFileName : string;

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  ShellAPI
  , ParamSoc
  , Windows
  , wCommuns
  , Mul
  , TntDBGrids
  , utilPGI
  ;

procedure TOF_YYJNALEVENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.OnUpdate ;
begin
  Inherited ;
  FListe_OnRowEnter(Self);
end ;

procedure TOF_YYJNALEVENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.OnArgument (S : String ) ;
var
  Arg : string;
begin
  Inherited ;
  VoirDoc            := TToolbarButton97(GetControl('VOIRDOC'));
  lFListe            := TFMul(Ecran).FListe;
  VoirDoc.OnClick    := VoirDoc_OnClick;
  lFListe.OnRowEnter := FListe_OnRowEnter;
  Arg := GetArgumentString(S, 'TYPEEVENT');
  if Arg <> '' then
    THValComboBox(GetControl('GEV_TYPEEVENT')).Value := Arg;
  THValComboBox(GetControl('GEV_TYPEEVENT')).Enabled := (Arg <> 'RGP');
  Arg := GetArgumentString(S, 'LABEL', False);
  if Arg <> '' then
    THEdit(GetControl('GEV_LIBELLE')).Text := Arg;
  
end ;

procedure TOF_YYJNALEVENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYJNALEVENT.VoirDoc_OnClick(Sender: TObject);
var
  Path : string;
begin
  Path := GetParamSocSecur('SO_BTEMPLFILEREF', '') + '\' + GetFileName;
  if not FileExists(Path) then
    PGIError(Format(TraduireMemoire('Le fichier %s n''existe pas.'), [Path]), Ecran.Caption)
  else
    ShellExecute(0, pchar('open'), pchar(Path), nil, nil, SW_SHOW);
end;

procedure TOF_YYJNALEVENT.FListe_OnRowEnter(Sender: TObject);
begin
  VoirDoc.Visible := (ExisteSQL('SELECT 1 FROM BFILES WHERE BF0_CODE = "' + GetBFilesKey + '"'));
end;

function TOF_YYJNALEVENT.GetBFilesKey: string;
begin
  Result := 'EVT;RGP;' + lFliste.DataSource.DataSet.FindField('GEV_NUMEVENT').AsString;
end;

function TOF_YYJNALEVENT.GetFileName: string;
var
  Sql : string;
  Qry : TQuery;
begin
  Sql := 'SELECT BF0_FILENAME'
       + ' FROM BFILES'
       + ' WHERE BF0_CODE = "' + GetBFilesKey + '"'
       ;
  Qry := OpenSQL(Sql, True);
  try
    Result := iif(not Qry.EOF, Qry.Fields[0].AsString, '');
  finally
    Ferme(Qry);
  end;
end;

Initialization
  registerclasses ( [ TOF_YYJNALEVENT ] ) ;
end.

