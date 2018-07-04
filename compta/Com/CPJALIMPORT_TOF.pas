unit CPJALIMPORT_TOF;

interface

Uses StdCtrls, Controls, Classes, forms, sysutils, windows,
     ComCtrls, HCtrls, HEnt1, HMsgBox, extCtrls,
{$IFDEF EAGLCLIENT}
      eMul, eFiche,maineagl,
{$ELSE}
      Mul, FE_main,  HDB, DB, uDbxDataSet,
{$ENDIF}
      HTB97, Dialogs, Hxlspas, UTOF, UTOB, HQry;

procedure CPLanceFiche_JournalDesimports(Args : String);

Type
  TOF_CPJALIMPORT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad             ; override ;
    private
{$IFDEF EAGLCLIENT}
    FListe    : THGrid;
    wBookMark : integer;
{$ELSE}
    FListe    : THDBGrid;
    wBookMark : tBookMark;
{$ENDIF}
    FEcran    : TFMul ;
    Q         : TQuery;
    procedure SupprOnClick(Sender: TObject);
    procedure SaveRow;
    procedure RestorRow;

{$IFDEF EAGLCLIENT}
    procedure ExtractionRapport(L : THGrid; Q : THQuery);
{$ELSE}
    procedure ExtractionRapport(L : THDBGrid; Q : THQuery);
{$ENDIF}
  end;

implementation

uses TntDBGrids;

procedure CPLanceFiche_JournalDesimports(Args : String);
begin
    AGlLanceFiche ('CP', 'CPJALIMPORT', '', '', Args);
end;

procedure TOF_CPJALIMPORT.OnArgument (S : String ) ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe'));
{$ELSE}
  FListe := THDBGrid(GetControl('FListe'));
{$ENDIF}
  FEcran := TFMul(Ecran) ;

  TButton(GetControl('BExport')).visible := TRUE;
  TToolBarButton97(GetControl('BArchive')).Hint := 'Extraction des rapports';
  TToolBarButton97(GetControl('BArchive')).OnClick        := SupprOnClick ;
end;

procedure TOF_CPJALIMPORT.OnLoad ;
begin
  Inherited ;

{$IFDEF EAGLCLIENT}
  FEcran.HMTrad.ResizeGridColumns(Fliste);
{$ELSE}
  FEcran.HMTrad.ResizeDBGridColumns(FListe) ;
{$ENDIF}
end;

procedure TOF_CPJALIMPORT.SupprOnClick(Sender: TObject);
var
    F : TFMul;
    Query : THQuery;
begin
    F := TFMul(Longint( Ecran ));
    if (F = Nil) then exit;

    if (FListe = Nil) then exit;

{$IFDEF EAGLCLIENT}
    FEcran.Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
    SaveRow;
    Query := F.Q;
    if (Query = Nil) then exit;
     if Query.EOF then
     begin
        PGIInfo ('Aucun élément sélectionné'); RestorRow; exit;
     end;

    ExtractionRapport(Fliste,Query);
    RestorRow;
end;

{$IFDEF EAGLCLIENT}
procedure TOF_CPJALIMPORT.ExtractionRapport(L : THGrid; Q : THQuery);
{$ELSE}
procedure TOF_CPJALIMPORT.ExtractionRapport(L : THDBGrid; Q : THQuery);
{$ENDIF}
var
    i : integer;
    FName, Fna : string;
    SD : TSaveDialog;
    s  : TMemoryStream;
begin
    // si rien de selectionné !
    if ((L.NbSelected = 0) and (not L.AllSelected)) then
    begin
        MessageAlerte('Aucun élément sélectionné.');
        exit;
    end;

    // message d'avertisement !
    SD := TSaveDialog.Create(Application);
    SD.DefaultExt := '*.xls';
    SD.Title := 'Extraction des rapports';
    SD.Filter := 'Fichier Texte (*.txt)|*.txt';
    if SD.execute then
      Fna := SD.FileName;
    SD.free;
    FName := ReadTokenPipe (Fna, '.');

    // destruction
    if (L.AllSelected) then // si tout ??
    begin
        Q.First;

        while (Not Q.EOF) do
        begin
            s := TMemoryStream.Create;
            TBlobField(Q.FindField('CPJ_RAPPORT')).SaveToStream(s);
            s.SaveToFile(FName+Q.FindField ('CPJ_CODE').asstring+FormatDateTime(Traduitdateformat('yyyymmdd'), Q.FindField ('CPJ_DATE').Asdatetime)+'.TXT');
            s.free;
            Q.Next;
        end;

        L.AllSelected:=False;
    end
    else // si pas tout ....
    begin
        for i := 0 to L.NbSelected-1 do
        begin
            L.GotoLeBookMark(i);
            {$IFDEF EAGLCLIENT}
            FEcran.Q.TQ.Seek(L.Row-1);
            {$ENDIF}
            s := TMemoryStream.Create;
            TBlobField(Q.FindField('CPJ_RAPPORT')).SaveToStream(s);
            s.SaveToFile(FName+Q.FindField ('CPJ_CODE').asstring+FormatDateTime(Traduitdateformat('yyyymmdd'), Q.FindField ('CPJ_DATE').Asdatetime)+'.TXT');
            s.free;
            // AVOIR ExecuteSQl ('DELETE FROM CPJALIMPORT WHERE CPJ_CODE="'+ Q.FindField ('CPJ_CODE').asstring +'" AND CPJ_DATE="'+
            //UsDateTime (Q.FindField ('CPJ_DATE').Asdatetime)+'"');
        end;

        L.ClearSelected;
    end;
    FEcran.BChercheClick(nil) ;

end;

procedure TOF_CPJALIMPORT.SaveRow;
begin
{$IFDEF EAGLCLIENT}
   Q := FEcran.Q.TQ;
   Q.Seek(FListe.Row-1);
   wBookMark := FListe.Row;
{$ELSE}
   Q := FEcran.Q;
   wBookMark := FListe.DataSource.DataSet.GetBookmark;
{$ENDIF}
end;

procedure TOF_CPJALIMPORT.RestorRow;
begin
{$IFDEF EAGLCLIENT}
      Q := FEcran.Q.TQ;
      FListe.Row := wBookMark;
      if FListe.Row >0 then
        Q.Seek( FListe.Row )
      else
        Q.First;
{$ELSE}
     Q.GotoBookmark(wBookMark);
{$ENDIF}
end;


Initialization
  registerclasses ( [ TOF_CPJALIMPORT ] ) ;
end.
