{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/03/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMODIFMARCHEST ()
Mots clefs ... : TOF;BTMODIFMARCHEST
*****************************************************************}
Unit BTMODIFMARCHEST_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
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
     Vierge,
     HTB97,
     UTOF ;

Type
  TOF_BTMODIFMARCHEST = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    StCell : string;
    GS,GBAS : THGrid;
    TOBDatas : TOB;
    TOBCumul : TOB;
    SG_MARCHE,SG_AVENANTS,SG_ARECEVOIR : Integer;
    procedure AfficheGrids;
    procedure constitueGrids;
    procedure ChargeTob;
    procedure SetEvents (State : boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSTopLeftChanged (Sender : TObject);
    procedure GSColWidthsChanged (Sender : TObject);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
    procedure AddAvenants (Sender : TObject);
  end ;

const
  stChamps = 'LIBELLE;MARCHE;AVENANTS;ARECEVOIR;';
Implementation

uses USpecifPOC;

procedure TOF_BTMODIFMARCHEST.OnNew ;
begin
  Inherited ;
  //
//         'BM1_AFFAIRE="'+LaTOB.GetString('AFFAIRE')+'" AND '+
//         'BM1_MARCHE="'+copy(LaTOB.GetString('CODEMARCHE'),1,8)+'" AND '+
  //
end ;

procedure TOF_BTMODIFMARCHEST.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFMARCHEST.OnUpdate ;
var II : Integer;
    TL : TOB;
    ecart : Double;
begin
  Inherited ;
  for II := 0 to TOBDatas.Detail.Count -1 do
  begin
    TL := TOBDatas.Detail[II];
    if (TL.GetString('TYPEDATA')='R02') and (TL.getDouble('ARECEVOIR') <> TL.GetDouble('BM1_MONTANT')) then
    begin
      if TL.GetDouble('BM1_MTPOC') <> 0 then TL.SetDouble('BM1_MTPOC', Arrondi(TL.GetDouble('BM1_MTPOC')* (TL.GetDouble('ARECEVOIR')/TL.GetDouble('BM1_MONTANT')),V_PGI.okDecV));
      if TL.GetDouble('BM1_MTPAIEDIR') <> 0 then TL.SetDouble('BM1_MTPAIEDIR', Arrondi(TL.GetDouble('BM1_MTPAIEDIR')* (TL.GetDouble('ARECEVOIR')/TL.GetDouble('BM1_MONTANT')),V_PGI.okDecV));
      ecart := ARRONDI(TL.getDouble('ARECEVOIR') - (TL.GetDouble('BM1_MTPOC')+ TL.GetDouble('BM1_MTPAIEDIR')),V_PGI.OkDecV);
      if ecart <> 0 then
      begin
        if TL.GetDouble('BM1_MTPOC') > TL.GetDouble('BM1_MTPAIEDIR') then
        begin
          TL.SetDouble('BM1_MTPOC',TL.GetDouble('BM1_MTPOC')+ecart);
        end else
        begin
          TL.SetDouble('BM1_MTPAIEDIR',TL.GetDouble('BM1_MTPAIEDIR')+ecart);
        end;
      end;
      TL.SetDouble('BM1_MONTANT',TL.getDouble('ARECEVOIR'));
      TL.UpdateDB(false); 
    end;
  end;
  LaTOB.SetString('MODIF','X');
end ;

procedure TOF_BTMODIFMARCHEST.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFMARCHEST.OnArgument (S : String ) ;
var arow,acol : Integer;
    cancel : boolean;
begin
  Inherited ;
  TOBDatas := TOB.Create ('BTMARCHESTDET',nil,-1);
  TOBCumul := TOB.Create ('LES CUMULS',nil,-1);
  TOBCumul.AddChampSupValeur('MARCHE',0);
  TOBCumul.AddChampSupValeur('AVENANTS',0);
  TOBCumul.AddChampSupValeur('ARECEVOIR',0);
  //
  GS := THGrid(GetControl('GS'));
  GBAS := THGrid(GetControl('GBAS'));
  constitueGrids;
  ChargeTob;
  AfficheGrids;
  SetEvents (True);
  Arow := 1;
  Acol := 1;
  GSCellEnter (Self,Acol,Arow,cancel);
  GBAS.Col := 0;
  GBAS.Row := 0;
  TToolbarButton97 (GetControl('Binsert')).onclick := AddAvenants; 
end ;

procedure TOF_BTMODIFMARCHEST.OnClose ;
begin
  TOBDatas.free;
  TOBCumul.free;
  Inherited ;
end ;

procedure TOF_BTMODIFMARCHEST.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFMARCHEST.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMODIFMARCHEST.AfficheGrids;
var II : Integer;
begin
  GS.RowCount := TOBDatas.Detail.count +1;
  for II := 0 to TOBDatas.Detail.count -1 do
  begin
    TOBDatas.detail[II].PutLigneGrid(GS,II+1,false,False,stChamps);
  end;
  TOBCumul.PutLigneGrid(GBAS,0,false,False,stChamps);
  TFVierge (Ecran).HMTrad.ResizeGridColumns(GS);
  for II := 0 to GS.ColCount -1 do
  begin
    GBAS.ColWidths [II] := GS.ColWidths [II];
  end;
end;

procedure TOF_BTMODIFMARCHEST.constitueGrids;
begin
  GS.cells[0,0] := '';
  GS.ColWidths [0] := 35 * GS.canvas.TextWidth('w') ;
  GS.ColEditables[0] := false;
  GS.ColLengths[0] := 35;
  //
  GS.cells[1,0] := 'Marché HT';
  GS.ColEditables[1] := false;
  GS.ColWidths [1] := 15 * GS.canvas.TextWidth('w') ;
  GS.ColLengths[1] := 15;
  GS.ColFormats[1] := '#0.00;#0.00; ;';
  GS.ColAligns[1] := taRightJustify;
  SG_MARCHE := 1;
  //
  GS.cells[2,0] := 'Avenants HT';
  GS.ColEditables[2] := false;
  GS.ColWidths [2] := 15 * GS.canvas.TextWidth('w') ;
  GS.ColLengths[2] := 15 ;
  GS.ColFormats[2] := '#0.00;#0.00; ;';
  GS.ColAligns[2] := taRightJustify;
  SG_AVENANTS := 2;
  //
  GS.cells[3,0] := 'A Régul.';
  GS.ColWidths [3] := 15 * GS.canvas.TextWidth('w') ;
  GS.ColLengths[3] := 15 ;
  GS.ColFormats[3] := '#0.00;#0.00; ;';
  GS.ColAligns[3] := taRightJustify;
  SG_ARECEVOIR := 3;
  //
  GBAS.cells[0,0] := 'Totaux';
  GBAS.ColLengths[0] := 35;
  //
  GBAS.ColFormats[1] := '#0.00;#0.00; ;';
  GBAS.ColAligns[1] := taRightJustify;
  //
  GBAS.ColFormats[2] := '#0.00;#0.00; ;';
  GBAS.ColAligns[2] := taRightJustify;
  //
  GBAS.ColFormats[3] := '#0.00;#0.00; ;';
  GBAS.ColAligns[3] := taRightJustify;
end;

procedure TOF_BTMODIFMARCHEST.ChargeTob;
var QQ : TQuery;
    II : Integer;
    SQL : string;
begin
  SQL := 'SELECT *,'+
         'BM1_LIBELLE AS LIBELLE,'+
         'TYPEDATA = '+
         'CASE '+
         'WHEN BM1_TYPE="001" THEN "001" '+
         'WHEN BM1_TYPE="002" AND SUBSTRING(BM1_SMARCHE,1,1)<>"R" THEN "002" '+
         'ELSE "R02" '+
         'END ,'+
         'IIF(BM1_TYPE="001",BM1_MONTANT,0) AS MARCHE,'+
         'IIF(BM1_TYPE="002" AND SUBSTRING(BM1_SMARCHE,1,1)<>"R",BM1_MONTANT,0) AS AVENANTS,'+
         'IIF(BM1_TYPE="002" AND SUBSTRING(BM1_SMARCHE,1,1)="R",BM1_MONTANT,0) AS ARECEVOIR '+
         'FROM BTMARCHESTDET '+
         'WHERE '+
         'BM1_AFFAIRE="'+LaTOB.GetString('AFFAIRE')+'" AND '+
         'BM1_MARCHE="'+copy(LaTOB.GetString('CODEMARCHE'),1,8)+'" AND '+
         'BM1_TYPE IN ("001","002")';
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBDatas.LoadDetailDB('BTMARCHESTDET','','',QQ,false);
  end;
  Ferme(QQ);
  for II := 0 to TOBDatas.Detail.Count -1 do
  begin
    TOBCumul.SetDouble('MARCHE',TOBCumul.GetDouble('MARCHE')+TOBDatas.Detail[II].GetDouble('MARCHE'));
    TOBCumul.SetDouble('AVENANTS',TOBCumul.GetDouble('AVENANTS')+TOBDatas.Detail[II].GetDouble('AVENANTS'));
    TOBCumul.SetDouble('ARECEVOIR',TOBCumul.GetDouble('ARECEVOIR')+TOBDatas.Detail[II].GetDouble('ARECEVOIR'));
  end;
end;

procedure TOF_BTMODIFMARCHEST.SetEvents (State : boolean);
begin
  if State then
  begin
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnTopLeftChanged  := GSTopLeftChanged;
    GS.OnColumnWidthsChanged := GSColWidthsChanged;
  end else
  begin
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
    GS.OnTopLeftChanged  := nil;
    GS.OnColumnWidthsChanged := nil;
  end;
end;

procedure TOF_BTMODIFMARCHEST.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if Not cancel then StCell := GS.cells[Acol,Arow];
end;

procedure TOF_BTMODIFMARCHEST.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TL : TOB;
begin
  //
  TL := TOBDatas.detail[Arow-1]; if TL=nil then Exit;
  if (Acol <> SG_ARECEVOIR) or (TL.GetString('TYPEDATA')<>'R02') then
  begin
    GS.Cells[ACol,Arow] := StCell;
    TL.PutLigneGrid(GS,Arow,false,False,stChamps);
  end else
  begin
    if Not cancel then
    begin
      TOBCumul.SetDouble('ARECEVOIR',TOBCumul.GetDouble('ARECEVOIR')-TL.GetDouble('ARECEVOIR'));
      TL.SetDouble('ARECEVOIR',Valeur(GS.cells[Acol,Arow]));
      TOBCumul.SetDouble('ARECEVOIR',TOBCumul.GetDouble('ARECEVOIR')+TL.GetDouble('ARECEVOIR'));
      TL.PutLigneGrid(GS,Arow,false,False,stChamps);
      TOBCumul.PutLigneGrid(GBAS,0,false,False,stChamps);
    end;
  end;
end;

procedure TOF_BTMODIFMARCHEST.GSTopLeftChanged(Sender: TObject);
var II : Integer;
begin
  GBAS.LeftCol  := GS.leftCol;
  for II := 0 to GS.ColCount -1 do
  begin
    GBAS.ColWidths [II] := GS.ColWidths [II];
  end;
  GBAS.invalidate;

end;



function TOF_BTMODIFMARCHEST.ZoneAccessible(ACol, ARow: Longint): boolean; { NEWPIECE }
var TOBL: TOB;
begin
  Result := True;
  if Arow > TOBDatas.detail.count then begin result := false; exit; end;
  TOBL := TOBDatas.detail[Arow-1];
  if TOBL = nil then Exit;
  if (Acol = 0) then BEGIN result := false; Exit; END;
//  if ((TOBL.GetString('TYPEDATA')<> 'R02') and ( Acol =  SG_ARECEVOIR)) then begin result := false; exit; end;
end;

procedure TOF_BTMODIFMARCHEST.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := GS.SynEnabled;
  GS.SynEnabled := False;
  Sens := -1;
  ChgLig := (GS.Row <> ARow);
  ChgSens := False;
  if GS.Row > ARow then Sens := 1 else if ((GS.Row = ARow) and (ACol <= GS.Col)) then Sens := 1;
  ACol := GS.Col;
  ARow := GS.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GS.RowCount;
      // ---
      if ((ACol = GS.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break;
      end;
      if ChgLig then
      begin
        ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        if ARow >= Lim then
        begin
          ARow := 1;
          Acol := 1;
        end else
        begin
          ACol := GS.FixedCols;
        end;
      end;
    end else
    begin
      if ((ACol = GS.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Arow := GS.rowCount;
          ChgLig := true;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS.ColCount;
        ChgLig := False;
      end;
      if ACol > GS.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GS.ColCount - 1;
      end;
    end;
  end;
  GS.SynEnabled := OldEna;
end;


procedure TOF_BTMODIFMARCHEST.GSColWidthsChanged(Sender: TObject);
var II : Integer;
begin
  for II := 0 to GS.ColCount -1 do
  begin
    GBAS.ColWidths [II] := GS.ColWidths [II];
  end;
end;

procedure TOF_BTMODIFMARCHEST.AddAvenants(Sender: TObject);
begin
  AppelMarcheST (LaTOB.GetString('AFFAIRE'),LaTOB.GetString('SOUSTRAIT'),Copy(LaTOB.GetString('CODEMARCHE'),1,8))
end;

Initialization
  registerclasses ( [ TOF_BTMODIFMARCHEST ] ) ;
end.
