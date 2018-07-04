{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion des cas particulers.
Suite ........ : Concerne les différents profils comme les bas salaires.
Suite ........ : En fait, cela correspond à une liste de profils traitant des cas 
Suite ........ : particuliers pour les exonérations
Mots clefs ... : PAIE;CASPARTICULIERS
*****************************************************************}
unit PartEtab;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HTB97,StdCtrls, Hctrls, Grids, UTOB, hmsgbox;

  PROCEDURE LancePartEtab (ETAB: STRING) ;

type
  TFCasPartEtab = class(TForm)
   Dock971: TDock97;
   ToolWindow971: TToolWindow97;
    PCas: THGrid;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
//    DataSource1: TDataSource;
//    Table1: TTable;
    THEME: THValComboBox;
    PPART: THValComboBox;
    BVALIDER: TToolbarButton97;
    BFERMER: TToolbarButton97;
    BHELP: TToolbarButton97;
    Msg: THMsgBox;
    procedure sauve;
    procedure FormShow(Sender: TObject);
    procedure BVALIDERClick(Sender: TObject);
    procedure PCasRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Déclarations privées }
    ETAB: String;
    Modifie : Boolean ;
  public
    { Déclarations publiques }
  end;

var
  FCasPartEtab: TFCasPartEtab;

implementation

{$R *.DFM}
PROCEDURE LancePartEtab (ETAB: STRING) ;
VAR
  X:  TFCasPartEtab ;
BEGIN
  X           := TFCasPartEtab.Create (Application) ;
  X.ETAB      := ETAB ;
  TRY
  X.ShowModal ;
  FINALLY
  X.Free;
  END ;
END ;

procedure TFCasPartEtab.FormShow(Sender: TObject);
var i,j : Integer;
    TPS : TOB;
    Profil,Themes : String;
begin
Modifie:=FALSE ;
if (PPART=Nil) or (PCAS=Nil) or (THEME=Nil) then exit ;
if (THEME<>Nil) and (PCAS<>Nil) then
   BEGIN
   PCAS.RowCount:=THEME.Items.Count+1 ;
   For i:=0 to THEME.Items.Count-1 do PCAS.Cells[0,i+1]:=THEME.Items[i] ;
   END ;
PPART.Plus:='' ;
TPS:=TOB.Create('profils speciaux',Nil,-1) ;
TPS.LoadDetailDB('PROFILSPECIAUX','"X";"'+ETAB+'"','',Nil,FALSE,FALSE) ;
For i:=0 to TPS.Detail.Count-1 do
    BEGIN
    Themes:=TPS.Detail[i].GetValue('PPS_THEMEPROFIL') ;
    PPART.Value:=TPS.Detail[i].GetValue('PPS_PROFIL') ;
    Profil:=PPART.Text ;
    for j:=0 to THEME.Items.Count-1 do
       if Themes=THEME.Values[j] then BEGIN PCas.Cells[1,j+1]:=Profil ; break ; END ;
    END ;
TPS.Free ;
end;

procedure TFCasPartEtab.BVALIDERClick(Sender: TObject);
begin
Sauve ;
ModalResult := mrOk ;
end;

procedure TFCasPartEtab.PCasRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
PPART.Plus:=THEME.Values[THGrid(Sender).Row-1] ;
Modifie := TRUE;
end;


procedure TFCasPartEtab.sauve;
var i : Integer;
    Profil : String;
begin
if Not Modifie then exit ;
if (PPART=Nil) or (PCAS=Nil) or (THEME=Nil) then exit ;
PPART.Plus:='' ;
ExecuteSQL('DELETE FROM PROFILSPECIAUX WHERE PPS_ETABSALARIE="X" AND PPS_CODE="'+ETAB+'"') ;
For i:=1 to PCAS.RowCount-1 do
    BEGIN
    PPART.Libelle:=PCAS.Cells[1,i] ;
    Profil:=PPART.Value ;
    if Profil<>'' then
     ExecuteSQL('INSERT INTO PROFILSPECIAUX (PPS_ETABSALARIE,PPS_CODE,PPS_PROFIL,PPS_THEMEPROFIL) '+
                ' VALUES '+
                ' ("X","'+ETAB+'","'+Profil+'","'+THEME.Values[i-1]+'")') ;
    END ;
end;

procedure TFCasPartEtab.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
Var Rep : integer ;
begin
if Modifie then Rep:=Msg.Execute(0,caption,'') else rep:=mrNo ;
Case Rep of
   mrYes    : Sauve ;
   mrNo     : Modifie:=FALSE ;
   mrCancel : CanClose:=FALSE ;
   END ;
end;

end.
