unit VisuExp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, DB, Hqry, Grids, DBGrids, HDB, StdCtrls, Buttons, ExtCtrls, HEnt1,Ent1, Hctrls,
  ADODB, udbxDataset ;


type
  TFVisuExp = class(TForm)
    PanelBouton: TPanel;
    BRechercher: THBitBtn;
    Panel1: TPanel;
    BImprimer: THBitBtn;
    BOuvrir: THBitBtn;
    BAnnuler: THBitBtn;
    BAide: THBitBtn;
    FListe: THDBGrid;
    QEcr: THQuery;
    SEcr: TDataSource;
    FindDialog: TFindDialog;
    HQuery1: THQuery;
    procedure FormCreate(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FirstFind : boolean;
    WMinX,WMinY : Integer ;
    Pref : String3 ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
  public
  end;

Procedure ListeExportes(StSQL : String ; Pref : String3) ;

implementation

uses
{$IFNDEF IMP}
     Saisie,
     eSaisBud,
{$ENDIF}
     PrintDBG,ParamDBG ;
{$R *.DFM}

Procedure ListeExportes(StSQL : String ; Pref : String3) ;
var FVisuExp:TFVisuExp ;
BEGIN
FVisuExp:=TFVisuExp.Create(Application) ;
try
  FVisuExp.Pref:=Pref ;
  FVisuExp.QEcr.SQL.Clear ;
  FVisuExp.QEcr.SQL.Add(StSQL) ;
//  FVisuExp.QEcr.SQL.Assign(StSQL) ;
  FVisuExp.ShowModal ;
  finally
  FVisuExp.Free ;
  END ;
Screen.Cursor:=SynCrDefault ;
END ;

procedure TFVisuExp.WMGetMinMaxInfo(var MSG: Tmessage);
begin
inherited;
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFVisuExp.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFVisuExp.FListeDblClick(Sender: TObject);
begin
{$IFNDEF IMP}
if (QEcr.Eof) And (QEcr.Bof) then Exit ;
if Pref='BE' then TrouveEtLanceSaisBud(QEcr,taConsult) 
             else TrouveEtLanceSaisie(QEcr,taConsult,'') ;
{$ENDIF}
end;

procedure TFVisuExp.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil, Caption,''); end;

procedure TFVisuExp.BRechercherClick(Sender: TObject);
begin FirstFind:=True; FindDialog.Execute ; end;

procedure TFVisuExp.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FirstFind); end;

procedure TFVisuExp.FormShow(Sender: TObject);
var LChamp : Array[0..9] of String ;
    i : integer ;
begin
LChamp[0]:='_JOURNAL' ; LChamp[1]:='_DATECOMPTABLE' ; LChamp[2]:='_NUMEROPIECE' ;
LChamp[3]:='_GENERAL' ; LChamp[4]:='_AUXILIAIRE' ; LChamp[5]:='_DEBIT' ;
LChamp[6]:='_CREDIT' ; LChamp[7]:='_NATUREPIECE' ; LChamp[8]:='_QUALIFPIECE' ;
if Pref='BE' then
  BEGIN
  LChamp[0]:='_BUDJAL' ;
  LChamp[3]:='_BUDGENE' ; LChamp[4]:='_BUDSECT' ;
  LChamp[4]:='_AXE' ;
  END ;
FListe.DataSource.DataSet.DisableControls ;
For i:=0 to FListe.Columns.Count-1 do
    BEGIN
    FListe.Columns[i].FieldName:=Pref+LChamp[i] ;
    if ((FListe.Columns[i].Field<>Nil) and (FListe.Columns[i].Field.DataType=ftfloat)) then TFloatField(FListe.Columns[i].Field).DisplayFormat:=StrFMask(V_PGI.OkDecV,'',True) ;
    END;
FListe.DataSource.DataSet.EnableControls ;
ChangeSQL(QEcr) ;
QEcr.Open ;
end;

end.
