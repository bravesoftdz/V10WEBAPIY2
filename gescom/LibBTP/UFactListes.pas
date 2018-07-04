unit UFactListes;

interface
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
       Hctrls,
       uTob,
       forms,
       sysutils,
       HEnt1,
       UentCommun,
       ParamSoc,
       windows,                                   
       graphics,
       Grids,
       UTOF ;
type

  TChps = class
    Nomchps : string;
    TypeChps : string;
    LibelleINI : string;
    Modifiable : boolean;
    TailleINI : integer;
    NbDecIni : integer;
    AlignINI : TCaligment;
    //
    Designation : string;
    TailleReele : integer;
    NbDec : integer;
    Align : TCaligment;
    SepMilliers : Boolean;
    LibComplet : boolean;
    Obligatoire : Boolean;
    visible : boolean;
    BlancSiNul : boolean;
    WithCumul : boolean;
    //
    CellCumulG : integer;
    RowCumulG : integer;
  public
    function GetSuffixe : string;
  end;

  TlistChps = class (Tlist)
  private
    function Add(AObject: TChps): Integer;
    function GetItems(Indice: integer): TChps;
    procedure SetItems(Indice: integer; const Value: TChps);
  public
    destructor destroy; override;
    property Items [Indice : integer] : TChps read GetItems write SetItems;
    function find (NomChps : string): TChps;
    procedure Clear; override;
  end;

  TListeSaisie = class (Tobject)
    private
      XX : TForm;
      fGS : THGrid;
      fGT : THgrid;
      FGTG : THGrid;
      fListChamps : TListChps;
      TS: TStringList;
      fCache : boolean;
      fIsExistCumul : boolean;
      fRecadreAuto : boolean;
      fActivated : boolean;
      fInresize : boolean;
      procedure DefiniListe;
      procedure VideGridGT;
      procedure RemplitGridGT (TOBPiece : TOB);
      procedure GTGetCellCanvas (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
      procedure RecupModifiable;
      procedure SetCadreAuto(const Value: boolean);
    public
      constructor Create (FF : TForm);
      destructor destroy; override;
      property IsExistCumul : boolean read fIsExistCumul;
      property Activate : boolean read fActivated write fActivated;
      property RecadreAuto : boolean read fRecadreAuto write SetCadreAuto;
      procedure AppliqueGrids;
      function SetListe (CodeListe : string) : boolean;
      procedure SetNomChamps (ListeSaisie : string);
      procedure AjusteGridTotal;
      procedure RemplitGridTotal (TOBpiece : TOB);
      procedure CacheGridTotal;
      procedure SetEtatGrids(cache: boolean);
      procedure FormResize;
      function ISCumulable (ACol : integer) : boolean;
      function ZoneModifiable (Acol : integer) : boolean;
      procedure SetColsModifiable;
      procedure DefiniTailleResize;
      procedure RecadreAutoDefaut;
      function ISZoneMontant(Acol: Integer): Boolean;
      function GetMtTotal(TOBL:TOB;Acol : integer) : double;
  end;

implementation

uses Facture, SaisUtil, EntGC
;

{ TlistChps }

function TlistChps.Add(AObject: TChps): Integer;
begin
	result := inherited add(Aobject);
end;

procedure TlistChps.Clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TChps(Items [Indice])<> nil then TChps(Items [Indice]).free;
    end;
  end;
  inherited;
end;

destructor TlistChps.destroy;
begin
  clear;
  inherited;
end;

function TlistChps.find(NomChps: string): TChps;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].Nomchps = NomChps then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistChps.GetItems(Indice: integer): TChps;
begin
  result := TChps (Inherited Items[Indice]);
end;

procedure TlistChps.SetItems(Indice: integer; const Value: TChps);
begin
  Inherited Items[Indice]:= Value;
end;

{ TListeSaisie }
procedure TListeSaisie.DefiniListe;
var NomChps,CurChps : string;
    TailleChps,CurTaill : string;
    LibChps,CurLib,TypeInfo,Libelle : string;
    DefChps,Curdef : string;
    TheItem : TChps;
    TailleINI : integer;
    LAlig,LSepMil,LnbDec,LLibC,LBlanc,LCUMUL,LLIbvisi,LOblig : string;
begin
  fListChamps.Clear;
  fListChamps.Pack;
  NomChps := TS.Strings[1];
  LibChps := TS.Strings[4];
  TailleChps := TS.Strings[5];
  DefChps := TS.Strings[6];
  Repeat
    CurChps := READTOKENST (NomChps);
    if (CurChps ='') or (CurChps=' ') then break;
    CurLib := READTOKENST (LibChps);
    Curdef := READTOKENST (DefChps);
    CurTaill := READTOKENST (TailleChps);
    // Decodage de la definition
    LAlig := Copy(CurDef,1,1);
    LSepMil := Copy(CurDef,2,1);
    LnbDec := Copy(CurDef,3,1);
    LOblig := Copy(CurDef,4,1);
    LLibC := Copy(CurDef,5,1);
    LLibVisi := Copy(CurDef,6,1);
    LBlanc := Copy(CurDef,7,1);
    LCUMUL :=Copy(CurDef,8,1);
    // ---
    GetInfoChamps (CurChps,TypeInfo,Libelle,TailleINI);
    //
    TheItem := TChps.Create;
    TheItem.TypeChps := TypeInfo;
    TheITem.Nomchps  := CurChps;
    TheITem.TailleINI := TailleINI;
    TheItem.Designation := CurLib;
    TheItem.TailleReele := StrToInt(CurTaill);
    TheItem.NbDec := StrToInt(LNbdec);
    if LAlig = 'G' then TheItem.Align := TalLeft
    else if LAlig = 'D' then TheItem.Align := TalRight
    else TheItem.Align := TalCenter;
    TheItem.SepMilliers := (LSepMil='/');
    TheItem.Obligatoire := (LOblig='O');
    TheItem.visible := (LLibVisi<>'X');
    TheItem.BlancSiNul := (LBlanc='X');
    TheItem.WithCumul := (LCUMUL='X');
    TheItem.CellCumulG := 0;
    TheItem.RowCumulG := 0;
    if (TheItem.TypeChps = 'DOUBLE') and (TheItem.WithCumul) then fIsExistCumul := true; 
    fListChamps.Add(TheItem);
    //
  Until CurChps='';
  // Initialisation
  RecupModifiable;
end;

constructor TListeSaisie.Create(FF: TForm);
begin
  fActivated := false;
  fInresize := false;
  fCache := false;
  fIsExistCumul := false;
  XX := nil;
  TS := TStringList.Create;
  fListChamps := TlistChps.Create;
  XX := TFFacture(FF);
  fGS := TFFacture(FF).GS;
  FGT := TFFActure(FF).GT;
  FGTG := TFFActure(FF).GTG;
  FGT.GetCellCanvas := GTGetCellCanvas;
end;

destructor TListeSaisie.destroy;
begin
  TS.free;
  fListChamps.free;
  inherited;
end;

function TListeSaisie.SetListe(CodeListe: string) : boolean;
var QQ : Tquery;
begin
  result := false;
  QQ := OpenSql ('SELECT * FROM LISTE WHERE LI_LISTE="'+CodeListe+'"',True,1,'',True);
  TRY
    if not QQ.eof then
    begin
      TS.SetText(PChar(QQ.FindField('LI_DATA').AsString));
      result := true;
    end;
  FINALLY
    ferme (QQ);
    if result then
    begin
      DefiniListe;
    end;
  END;
end;


procedure TListeSaisie.AppliqueGrids;
var II,i : integer;
    StDecM: string;
begin
  //
  FGTG.FixedColor := clActiveCaption;
  FGTG.colcount := 9;
  FGTG.Rowcount := 4;
  FGTG.Cells [1,0] := 'Fourn.';
  FGTG.Cells [2,0] := 'Mo Int.';
  FGTG.Cells [3,0] := 'S.Trait';
  FGTG.Cells [4,0] := 'Intérim.';
  FGTG.Cells [5,0] := 'Location';
  FGTG.Cells [6,0] := 'Matériel';
  FGTG.Cells [7,0] := 'Outillage';
  FGTG.Cells [8,0] := 'Autres';
  //
  FGTG.Cells [0,1] := 'PA';
  FGTG.Cells [0,2] := 'PR';
  FGTG.Cells [0,3] := 'PV';
  //
  StDecM := '#,##0.';
  if TFFACture(XX).DEV.Decimale > 0 then
  begin
    for i := 1 to TFFACture(XX).DEV.Decimale - 1 do
    begin
      StDecM := StDecM + '0';
    end;
    StDecM := StDecM + '0';
  end;
  //
  for II := 1 to 8 do
  begin
    FGTG.ColFormats[II] := StDecM+';'+StDecM+';;';
    FGTG.ColTypes [II] := 'R';
    FGTG.ColAligns  [II] := taRightJustify;
    FGTG.ColWidths [II] := FGTG.Canvas.TextWidth('W') * 12; // Taille d'affichage
    FGTG.ColLengths [II] := 12; // Taille d'affichage
  end;
end;

procedure TListeSaisie.AjusteGridTotal;
var II : integer;
    StDecM: string;
begin
  //
  FGT.left := fGS.left;
  FGT.Width  := FGS.Width - GetSystemMetrics(SM_CYVSCROLL);
  FGT.Rowcount := 1;
  FGT.ColCount := fGS.ColCount;
  StDecM := '#,##0.';
  if TFFACture(XX).DEV.Decimale > 0 then
  begin
    for II := 1 to TFFACture(XX).DEV.Decimale - 1 do
    begin
      StDecM := StDecM + '0';
    end;
    StDecM := StDecM + '0';
  end;
  //
  for II := 0 to fGS.Colcount -1 do
  begin
    if II= 0 then FGT.ColWidths [II] := FGS.ColWidths [II] - 1
             else FGT.ColWidths [II] := FGS.ColWidths [II] + 1;
    FGT.ColLengths [II] := FGS.ColLengths [II];
    FGT.ColAligns  [II] := taRightJustify;
  end;
end;

procedure TListeSaisie.SetNomChamps(ListeSaisie: string);
var II : Integer;
    LesChamps,LeChamp : string;
begin
  II := 0;
  LesChamps := ListeSaisie;
  repeat
    LeChamp := READTOKENST(LesChamps);
    if LeChamp ='' then break;
    fListChamps.Items[II].Nomchps := LeChamp;
    inc(II);
  until LeChamp = '';
end;

procedure TListeSaisie.RemplitGridTotal(TOBpiece : TOB);
begin
  RemplitGridGT (TOBPiece);
  if not fCache then
  begin
  FGTG.Cells [1,1] := strf00(TOBPiece.GetDouble('GP_TOTALFOUPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [1,2] := strf00(TOBPiece.GetDouble('GP_TOTALFOUPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [1,3] := strf00(TOBPiece.GetDouble('GP_TOTALFOUPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [2,1] := strf00(TOBPiece.GetDouble('GP_TOTALMOPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [2,2] := strf00(TOBPiece.GetDouble('GP_TOTALMOPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [2,3] := strf00(TOBPiece.GetDouble('GP_TOTALMOPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [3,1] := strf00(TOBPiece.GetDouble('GP_TOTALSTPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [3,2] := strf00(TOBPiece.GetDouble('GP_TOTALSTPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [3,3] := strf00(TOBPiece.GetDouble('GP_TOTALSTPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [4,1] := strf00(TOBPiece.GetDouble('GP_TOTALINTPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [4,2] := strf00(TOBPiece.GetDouble('GP_TOTALINTPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [4,3] := strf00(TOBPiece.GetDouble('GP_TOTALINTPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [5,1] := strf00(TOBPiece.GetDouble('GP_TOTALLOCPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [5,2] := strf00(TOBPiece.GetDouble('GP_TOTALLOCPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [5,3] := strf00(TOBPiece.GetDouble('GP_TOTALLOCPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [6,1] := strf00(TOBPiece.GetDouble('GP_TOTALMATPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [6,2] := strf00(TOBPiece.GetDouble('GP_TOTALMATPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [6,3] := strf00(TOBPiece.GetDouble('GP_TOTALMATPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [7,1] := strf00(TOBPiece.GetDouble('GP_TOTALOUTPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [7,2] := strf00(TOBPiece.GetDouble('GP_TOTALOUTPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [7,3] := strf00(TOBPiece.GetDouble('GP_TOTALOUTPV'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [8,1] := strf00(TOBPiece.GetDouble('GP_TOTALAUTPA'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [8,2] := strf00(TOBPiece.GetDouble('GP_TOTALAUTPR'),TFFActure(XX).DEV.Decimale);
  FGTG.Cells [8,3] := strf00(TOBPiece.GetDouble('GP_TOTALAUTPV'),TFFActure(XX).DEV.Decimale);
  end else
  begin
    FGTG.Cells [1,1] := '';
    FGTG.Cells [1,2] := '';
    FGTG.Cells [1,3] := '';
    FGTG.Cells [2,1] := '';
    FGTG.Cells [2,2] := '';
    FGTG.Cells [2,3] := '';
    FGTG.Cells [3,1] := '';
    FGTG.Cells [3,2] := '';
    FGTG.Cells [3,3] := '';
    FGTG.Cells [4,1] := '';
    FGTG.Cells [4,2] := '';
    FGTG.Cells [4,3] := '';
    FGTG.Cells [5,1] := '';
    FGTG.Cells [5,2] := '';
    FGTG.Cells [5,3] := '';
    FGTG.Cells [6,1] := '';
    FGTG.Cells [6,2] := '';
    FGTG.Cells [6,3] := '';
    FGTG.Cells [7,1] := '';
    FGTG.Cells [7,2] := '';
    FGTG.Cells [7,3] := '';
    FGTG.Cells [8,1] := '';
    FGTG.Cells [8,2] := '';
    FGTG.Cells [8,3] := '';
  end;
  TFFacture(XX).HMTrad.ResizeGridColumns(FGTG);
  FGTG.CacheEdit;
end;

procedure TListeSaisie.CacheGridTotal;
var II,JJ : integer;
begin
  VideGridGT ;
  for II := 1 to 8 do
  begin
    for JJ := 1 to 3 do
    begin
      FGTG.Cells [II,JJ] := '';
    end;
  end;
  TFFacture(XX).HMTrad.ResizeGridColumns(FGTG);
  FGTG.enabled;
  FGTG.CacheEdit;
end;

procedure TListeSaisie.SetEtatGrids (cache : boolean);
begin
  fCache := cache;
  if not fCache then RemplitGridGT(TFFACTure(XX).LaPieceCourante)
                else CacheGridTotal;
end;

procedure TListeSaisie.VideGridGT;
var II : integer;
begin
  for II := 0 to fListChamps.count -1 do
  begin
    if fListChamps.Items[II].WithCumul then
    begin
      FGT.Cells [II,0] := '';
    end;
  end;
end;

procedure TListeSaisie.RemplitGridGT(TOBPiece : TOB);
var II : integer;
    TheChamps,GPChamps : string;
    TheValeur : string;
begin
  if fCache then exit;
  for II := 0 to fListChamps.count -1 do
  begin
    GPChamps := '';
    if fListChamps.Items[II].WithCumul then
    begin
      TheChamps := fListChamps.Items[II].Nomchps;
      if pos('_',TheChamps) <0  then continue;
      // Cas particulier .. a la con
      if ((TheChamps = 'GL_MONTANTHTDEV') or (TheChamps = 'GL_MONTANTHT')) and (TOBPiece.GetBoolean('GP_PIECEENPA')) then
      begin
        GPChamps := 'GP_MONTANTPA';
      end else if TheChamps = 'GL_HEURE' then
      begin
        TheChamps := 'GL_TOTALHEURE';
        GPChamps := 'GP_'+copy(TheCHamps,pos('_',TheChamps)+1,length(TheCHamps));
      end else if copy(TheChamps,1,10) = 'GL_MONTANT' then
      begin
        TheChamps := 'GL_TOTAL'+copy(TheCHamps,11,length(TheChamps));
        GPChamps := 'GP_'+copy(TheCHamps,pos('_',TheChamps)+1,length(TheCHamps));
      end else if copy (theChamps,1,6)='GLC_MT' then
      begin
        GPChamps := GetChampsTot(TheChamps);
      end else if (pos(TheChamps,ZONETOTAL)>0) then
      begin
        GPChamps := GetChampsTot(TheChamps);
      end;
      if TOBPiece.FieldExists(GPChamps) then
      begin
        TheValeur := strf00(TOBPiece.GetDouble(GPChamps),TFFActure(XX).DEV.Decimale);
        FGT.Cells [II,0] := TheValeur;
      end;
    end;
  end;
end;

procedure TListeSaisie.GTGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);     
begin
  if ACol < fGT.FixedCols then Exit;
  if Arow < fGT.Fixedrows then Exit;
  Canvas.Font.Style := Canvas.Font.Style + [fsbold];
end;

procedure TListeSaisie.FormResize;
begin
  AjusteGridTotal;
  TFFacture(XX).HMTrad.ResizeGridColumns(FGTG);
  if fRecadreAuto then TFFacture(XX).HMTrad.ResizeGridColumns(FGS);
  FGTG.CacheEdit;
end;

function TListeSaisie.ISZoneMontant (Acol : Integer) : Boolean;
begin
  result := false;
  if (Acol < 0) or (fListChamps <> nil) and (Acol >= fListChamps.Count) then exit;
  result := (Copy(fListChamps.Items[ACol].Nomchps,1,6) = 'GLC_MT') or (pos(fListChamps.Items[ACol].Nomchps,ZONETOTAL)>0);
end;

function TListeSaisie.ISCumulable(ACol: integer): boolean;
begin
  result := false;
  if (Acol < 0) or (fListChamps <> nil) and (Acol >= fListChamps.Count) then exit;
  result := (fListChamps.Items[ACol].TypeChps = 'DOUBLE') and (fListChamps.Items[ACol].WithCumul);
end;

procedure TListeSaisie.SetColsModifiable;
var II : integer;
begin
  for II := 0 to fListChamps.count -1 do
  begin
    if not fListChamps.Items[II].Modifiable then
    begin
      FGS.ColEditables  [II] := false;
      FGS.ColLengths  [II] := 0;
    end;
  end;
end;

procedure TListeSaisie.RecupModifiable;
var UneListe,UnRecord,UR,UnChamps,Modif : string;
    UnElt : TChps;
begin
  UneListe := LISTCHPS;
  if VH_GC.BTCODESPECIF = '001' then UneListe := UneListe + ';SUMTOTALTS|N;MTTRANSFERT|N;NUMTRANSFERT|N;'; // ajout colonne total des avenants sur ligne pour POC
  repeat
    UnRecord := READTOKENST(UneListe);
    if Unrecord <> '' then
    begin
      UR := Unrecord;
      UnChamps := READTOKENPipe(UR,'|');
      Modif := 'O';
      if UR <> '' then Modif := UR;
      //
      UnELt := fListChamps.find(UnChamps);
      if UnElt <> nil then  UnElt.Modifiable := (Modif='O');
    end;
  until UnRecord = '';
end;

function TListeSaisie.ZoneModifiable(Acol: integer): boolean;
begin
  result := false;
  if (Acol < 0) or (Acol >= fListChamps.Count) then exit;
  result := (fListChamps.Items[ACol].Modifiable);
end;

procedure TListeSaisie.DefiniTailleResize;
begin
  //
end;

procedure TListeSaisie.SetCadreAuto(const Value: boolean);
begin
  if not Value then
  begin
    TFFacture(XX).HMTrad.LockedCtrls.Add('GS')
  end else
  begin
    TFFacture(XX).HMTrad.LockedCtrls.Clear;
  end;
  fRecadreAuto := Value;
end;

procedure TListeSaisie.RecadreAutoDefaut;
var ThePiece : string;
begin
  ThePiece := TFFacture(XX).CleDoc.NaturePiece;
  RecadreAuto := (GetInfoParpiece (ThePiece,'GPP_TAILLAGEAUTO')='X');
end;

function TListeSaisie.GetMtTotal(TOBL: TOB; Acol: integer): double;
var NomChps,NomTot : string;
begin
  result := 0;
  if (Acol < 0) or (fListChamps <> nil) and (Acol >= fListChamps.Count) then exit;
  NomChps := fListChamps.Items[ACol].Nomchps;
  if (Copy(NomChps,1,6)<> 'GLC_MT') and (Pos(NomChps,ZONETOTAL)<0) then Exit;
//  NomTot := 'TOTAL'+copy(NomChps,7,Length(Nomchps)-6);
  Result := TOBL.GetDouble(NomChps);
end;

{ TChps }

function TChps.GetSuffixe: string;
var II : integer;
begin
  II := Pos('_',Nomchps);
  if II > 0 then result := copy (Nomchps,II,SIzeof(Nomchps));
end;

end.
