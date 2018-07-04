{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /    
Description .. : Unité de paramétrage des domaines
Mots clefs ... : 
*****************************************************************}

unit UPDomaine;

interface

uses
  Windows, Messages, Classes, SysUtils, Graphics, Controls, StdCtrls, Forms,
  Dialogs, DBCtrls, DB, DBGrids,
  Grids, ExtCtrls, HPanel, UIUtil,
  hmsgbox, HSysMenu, HEnt1, menus,
  HCtrls,
  uDbxDataSet, Variants, ADODB,
  {$IFDEF EAGLCLIENT}
  UTob, UScriptTob,
  {$ENDIF}
  Uscript;

type
  TFDomaine = class(TForm)
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    Panel2: TPanel;
    HMTrad: THSystemMenu;
    ADOTable1: TADOTable;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { déclarations privées }
  public
    { déclarations publiques }
  end;

var
  FDomaine: TFDomaine;
  
procedure ParamDomaine(PP : THPanel);
procedure ChargementComboDomaine (Domaine : THVALComboBox; oktous : Boolean=TRUE);
function RendDomaine (Libelle : string): string;
function RendLibelleDomaine (Domaine : string): string;
procedure PurgePopup( PP : TPopupMenu ) ;

implementation

uses UDMIMP;

{$R *.DFM}

procedure ParamDomaine(PP : THPanel);
var XX : TFDomaine ;
BEGIN
XX:=TFDomaine.Create(Application);
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free ;
    End ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
end;

procedure TFDomaine.FormCreate(Sender: TObject);
begin
      ADOTable1.Active := TRUE;
end;

procedure ChargementComboDomaine (Domaine : THVALComboBox; oktous : Boolean=TRUE);
var
  Imptable: TADOTable;
  Vales, Libes   : HTstringList;
  dom       : string;
begin
  Vales := HTStringList.Create;
  Libes := HTStringList.Create;
  if not Assigned(dmImport) then
    dmImport := TdmImport.Create(Application);
  ImpTable := TADOTable.Create(Application);
  ImpTable.Connection := DMImport.DBGLOBAL;
  ImpTable := DMImport.GzImpDomaine;
  if oktous then  Vales.add ('<Tous>');
  with ImpTable do
  begin
    if Active then
      Close;
    Open;
    First;
    while not Eof do
    begin
      if  (ParamCount > 0) and (GetInfoVHCX.Domaine <> '') then
        if  ImpTable.FieldByName('Domaine').asstring <>  GetInfoVHCX.Domaine then
        begin
          Next;
          continue;
        end;

      dom := ImpTable.FieldByName('Libelle').asstring;
      if (Vales.IndexOf(dom) < 0) then
      begin
        Libes.add(dom);
        Vales.add(ImpTable.FieldByName('Domaine').asstring);
      end;
      Next;
    end;
    Domaine.Items := Libes;
    Domaine.Values := Vales;
    Vales.clear;
    Vales.Free;
    Libes.clear;
    Libes.free;
    Close;
  end;
end;

{$IFNDEF CISXPGI}
function RendDomaine (Libelle : string): string;
var
  Imptable: TADOTable;
  dom     : string;
begin
  if not Assigned(dmImport) then
    dmImport := TdmImport.Create(Application);
  ImpTable := DMImport.GzImpDomaine;
  with ImpTable do
  begin
    if Active then
      Close;
    Open;
    First;
    while not Eof do
    begin
      if Libelle = ImpTable.FieldByName('Libelle').asstring then
      begin
           dom := ImpTable.FieldByName('Domaine').asstring;
           break;
      end;
      Next;
    end;
    Close;
    Result := dom;
end;

end;

function RendLibelleDomaine (Domaine : string): string;
var
  Imptable: TADOTable;
  lib     : string;
begin
  if not Assigned(dmImport) then
    dmImport := TdmImport.Create(Application);
  ImpTable := DMImport.GzImpDomaine;
  with ImpTable do
  begin
    if Active then
      Close;
    Open;
    First;
    while not Eof do
    begin
      if Domaine = ImpTable.FieldByName('Domaine').asstring then
      begin
           lib := ImpTable.FieldByName('Libelle').asstring;
           break;
      end;
      Next;
    end;
    Close;
    Result := lib;
end;

end;

{$ELSE}
function RendDomaine (Libelle : string): string;
var
  Imptable: TQuery;
begin
  ImpTable := OpenSQL ('SELECT CO_LIBELLE, CO_CODE FROM COMMUN Where CO_TYPE="CSD" and CO_LIBELLE="'+
  Libelle +'"', TRUE);
  if not ImpTable.EOF then
           Result := ImpTable.FindField('CO_CODE').asstring;
  Ferme(ImpTable);
end;

function RendLibelleDomaine (Domaine : string): string;
var
  Imptable: TQuery;
begin
  ImpTable := OpenSQL ('SELECT CO_LIBELLE, CO_CODE FROM COMMUN Where CO_TYPE="CSD" and CO_CODE="'+
  Domaine +'"', TRUE);
  if not ImpTable.EOF then
           Result := ImpTable.FindField('CO_LIBELLE').asstring;
  Ferme(ImpTable);
end;
{$ENDIF}


procedure TFDomaine.FormShow(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
        HMTrad.ResizeDBGridColumns(DBGrid1);
        CentreDBGrid(DBGrid1);
{$ENDIF}
end;

procedure TFDomaine.FormClose(Sender: TObject; var Action: TCloseAction);
begin
ADOTable1.Close;
end;

procedure PurgePopup( PP : TPopupMenu ) ;
Var M,N : TMenuItem ;
BEGIN
if PP=Nil then Exit ;
if PP.Items.Count<=0 then Exit ;
While PP.Items.Count>0 do
   BEGIN
   M:=PP.Items[0] ;
   While M.Count>0 do BEGIN N:=M.Items[0] ; M.Remove(N) ; N.Free ; END ;
   PP.Items.Remove(M) ; M.Free ;
   END ;
END ;


end.