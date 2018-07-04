unit UFileAssoc;

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
     FE_Main,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Vierge,
     UentCommun,
     HTB97,
     Windows,
     AglInit ;

const
  TheChamps = ';BF0_LIBELLE;';

Type
  TFileAssoFunc = (TfafNone,TfafInterv,TfafAffaire,TFafClient,TfafDocument,TfafEvt);

  R_EVT = RECORD
            TypeEvt : string;
            NumEvt  : string ;
          END ;

  TOF_BTASSOCIEDOCS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    EmplacementStockage : string;
    GS : THGrid;
    TOBFiles : TOB;
    procedure GetFiles;
    procedure PrepareLaGrid;
    procedure AfficheLaGrid;
    procedure SetEvents(State : boolean);
    procedure GSEnter (Sender: TObject);
    procedure PrepareLibelle;
    procedure BInsertClick(Sender : TObject);
    procedure BDeleteClick(Sender : TObject);
    procedure BVoirDocClick(Sender: TObject);
  end ;

procedure StockeDocumentLie (FonctionSto : TFileAssoFunc;TOBDATA : TOB; FileName : string);
function ConstitueReferenceStockage (FonctionSto : TFileAssoFunc; TOBDATA : TOB) : string;
procedure GestionDocumentsLies (FonctionSto : TFileAssoFunc; TOBDATA : TOB; Action : TactionFiche);
function TypeGestionFictoText (FonctionSto : TFileAssoFunc) : string;
function LibelleAction(TOBDATA : TOB) : string;

procedure DecodeRefCliFiles(TOBData : TOB; var CodeClient : string);
procedure DecodeRefIntFiles(TOBData : TOB; var CodeAppel : string);
procedure DecodeRefDocFiles(TOBData : TOB; var Cledoc : r_cledoc);
procedure DecodeRefEvtFiles(TOBData : TOB; var CleEvt : r_Evt);
procedure DecodeRefAFFFiles(TOBData : TOB; var CodeAffaire : String);
function StoreFileInRef (RefPath,FileName : string) : string;

Implementation

uses Graphics,CalcOLEGenericBTP,UtilFichiers,Paramsoc;


function StoreFileInRef (RefPath,FileName : string) : string;
var ExtSrc,RacineFicDest,FicSrc,NomFicDest : string;
    SrcFile,DestFile : PWideChar;
    LnSrc,LnDest : Integer;
    OKOK : Boolean;
begin
  result := '';
  ExtSrc := ExtractFileExt(FileName);
  RacineFicDest := AglGetGuid;
  RacineFicDest := StringReplace(RacineFicDest,'-','',[rfreplaceAll]);
  NomFicDest := IncludeTrailingBackslash(RefPath)+RacineFicDest+ExtSrc;
  FicSrc := FileName;
  //
  LnSrc := (2*Length(FicSrc))+1;
  LnDest := (2*Length(NomFicDest))+1;
  //
  SrcFile := AllocMem(LnSrc);
  DestFile := AllocMem(LnDest);
  TRY
    StringToWideChar(FicSrc,SrcFile,lnSrc);
    StringToWideChar(NomFicDest,DestFile,LnDest);
  //
    OKOK := CopyFileW(SrcFile,DestFile,false);
    if not OKOK then
    begin
      RaiseLastOSError;
    end else
    begin
      result := ExtractFileName(NomFicDest);
    end;
  FINALLY
    FreeMem(SrcFile);
    FreeMem(DestFile);
  end;
end;


procedure StockeDocumentLie (FonctionSto : TFileAssoFunc;TOBDATA : TOB; FileName : string);
var CodeSto,SToFile,EmplacementStockage : string;
    TT : TOB;
begin
  if FileName ='' then Exit;
  EmplacementStockage := GetParamSocSecur ('SO_BTEMPLFILEREF','');
  if EmplacementStockage = '' then
  begin
    PGIInfo('Veuillez renseigner le paramètre de stockage des fichiers de référence');
    Exit;
  end;
  //
  CodeSto := ConstitueReferenceStockage (FonctionSto,TOBDATA);
  if CodeSto = '' then Exit;
  StoFile := StoreFileInRef(EmplacementStockage,FileName);
  if SToFile = '' then exit;
  TT := TOB.Create ('BFILES',nil,-1);
  TRY
    TT.SetString('BF0_CODE', CodeSto);
    TT.SetString('BF0_FILENAME', StoFile);
    TT.SetAllModifie(true);
    TT.InsertDB(nil);
  FINALLY
    TT.Free;
  END;
end;


procedure DecodeRefCliFiles(TOBData : TOB; var CodeClient : string);
var TheRef,UneInfo : string;
begin
  CodeClient := '';
  TheRef := TOBData.getString('CODE');
  if TheRef = '' then exit;
  UneInfo := READTOKENST(TheRef); // pour Enlever le "APP"
  CodeClient := READTOKENST(TheRef);
end;

procedure DecodeRefIntFiles(TOBData : TOB; var CodeAppel : string);
var TheRef,UneInfo : string;
begin
  CodeAppel := '';
  TheRef := TOBData.getString('CODE');
  if TheRef = '' then exit;
  UneInfo := READTOKENST(TheRef); // pour Enlever le "APP"
  CodeAppel := READTOKENST(TheRef);
end;

procedure DecodeRefAFFFiles(TOBData : TOB; var CodeAffaire : String);
var TheRef,UneInfo : string;
begin
  CodeAffaire := '';
  TheRef := TOBData.getString('CODE');
  if TheRef = '' then exit;
  UneInfo := READTOKENST(TheRef); // pour Enlever le "AFF"
  CodeAffaire := READTOKENST(TheRef);
end;

procedure DecodeRefDocFiles(TOBData : TOB; var Cledoc : r_cledoc);
var TheRef,UneInfo : string;
begin
  fillchar(cledoc,SizeOf(cledoc),#0);
  TheRef := TOBData.getString('CODE');
  if TheRef = '' then exit;
  UneInfo := READTOKENST(TheRef); // pour enlever le "DOC"
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then cledoc.NaturePiece := UneInfo;
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then cledoc.Souche := UneInfo;
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then cledoc.NumeroPiece := StrToInt(UneInfo);
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then cledoc.Indice := StrToInt(UneInfo);
end;

procedure DecodeRefEvtFiles(TOBData : TOB; var CleEvt : r_Evt);
var TheRef,UneInfo : string;
begin
  fillchar(CleEvt,SizeOf(CleEvt),#0);
  TheRef := TOBData.getString('CODE');
  if TheRef = '' then exit;
  UneInfo := READTOKENST(TheRef); // pour Enlever le "EVT"
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then CleEvt.TypeEvt := UneInfo;
  UneInfo := READTOKENST(TheRef);
  if UneInfo <> '' then CleEvt.NumEvt := UneInfo;
end;

procedure GestionDocumentsLies (FonctionSto : TFileAssoFunc; TOBDATA : TOB; Action : TactionFiche);
begin
  TheTOB := TOBDATA;
  AGLLanceFiche('BTP','BTASSOCIEDOCS','','','ACTION='+ActionToString(Action) );
  TheTOB := nil;
end;

function LibelleAction(TOBDATA : TOB) : string;
var cledoc : r_cledoc;
    cleEvt : R_EVT;
    code : string;
begin
  if TOBDATA.GetString('TYPE')='APP' then
  begin
    DecodeRefIntFiles (TOBDATA,Code);
    Result := 'Appel : '+  BTPCodeAffaireAffiche(Code);
  end else if TOBDATA.GetString('TYPE')='AFF' then
  begin
    DecodeRefAffFiles (TOBDATA,Code);
    Result := 'Affaire : '+  BTPCodeAffaireAffiche(Code);
  end else if TOBDATA.GetString('TYPE')='CLI' then
  begin
    DecodeRefCliFiles (TOBDATA,Code);
    Result := 'Client : '+  TOBDATA.GetString('CODE')
  end else if TOBDATA.GetString('TYPE')='DOC' then
  begin
    DecodeRefDocFiles (TOBDATA,Cledoc);
    Result := RechDom('GCNATUREPIECEG',Cledoc.NaturePiece,false)+' N° : '+InttoStr(cledoc.NumeroPiece);
  end else if TOBDATA.GetString('TYPE')='EVT' then
  begin
    DecodeRefEvtFiles (TOBDATA,cleEvt);
    Result := 'Evenement : '+cleEvt.TypeEvt+' N° : '+CleEvt.NumEvt;
  end;
end;


function TypeGestionFictoText (FonctionSto : TFileAssoFunc) : string;
begin
  Result := '';
  if FonctionSto = TfafNone then exit;
  if FonctionSto = TfafInterv then
  begin
    Result := 'APP';
  end else if FonctionSto = TfafAffaire then
  begin
    Result := 'AFF';
  end else if FonctionSto = TFafClient then
  begin
    Result := 'CLI';
  end else if FonctionSto = TfafDocument then
  begin
    Result := 'DOC;';
  end else if FonctionSto = TfafEvt then
  begin
    Result := 'EVT;';
  end;
end;

function ConstitueReferenceStockage (FonctionSto : TFileAssoFunc; TOBDATA : TOB) : string;
begin
  Result := '';
  if FonctionSto = TfafNone then exit;
  if FonctionSto = TfafInterv then
  begin
    if (TOBDATA.getString('AFF_AFFAIRE')='') then Exit;
    Result := 'APP;'+TOBDATA.getString('AFF_AFFAIRE');
  end else if FonctionSto = TfafAffaire then
  begin
    if (TOBDATA.getString('AFF_AFFAIRE')='') then Exit;
    Result := 'AFF;'+TOBDATA.getString('AFF_AFFAIRE');
  end else if FonctionSto = TFafClient then
  begin
    if (TOBDATA.getString('T_TIERS')='') then Exit;
    Result := 'CLI;'+TOBDATA.getString('T_TIERS');
  end else if FonctionSto = TfafDocument then
  begin
    if (TOBDATA.getInteger('GP_NUMERO')=0) or (TOBDATA.getString('GP_NATUREPIECEG')='') or (TOBDATA.getString('GP_SOUCHE')='') then Exit;
    Result := 'DOC;'+TOBDATA.getString('GP_NATUREPIECEG')+';'+TOBDATA.getString('GP_SOUCHE')+';'+TOBDATA.getString('GP_NUMERO')+';'+TOBDATA.getString('GP_INDICEG');
  end else if FonctionSto = TfafEvt then
  begin
    if (TOBDATA.getString('GEV_TYPEEVENT')='') or (TOBDATA.getInteger('GEV_NUMEVENT')=0) then Exit;
    Result := 'EVT;'+TOBDATA.getString('GEV_TYPEEVENT')+';'+TOBDATA.getString('GEV_NUMEVENT');
  end;
end;


procedure TOF_BTASSOCIEDOCS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.OnUpdate ;
begin
  Inherited ;
  BEGINTRANS;
  try
    ExecuteSQL('DELETE FROM BFILES WHERE BF0_CODE="'+LaTOB.GetString('CODE')+'"');
    TOBFiles.SetAllModifie(true);
    TOBFiles.InsertDB(nil);
    COMMITTRANS;
  except
    ROLLBACK;
  end;
end ;

procedure TOF_BTASSOCIEDOCS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.OnArgument (S : String ) ;
begin
  Inherited ;
  EmplacementStockage := GetParamSocSecur ('SO_BTEMPLFILEREF','');
  GS := THGrid(GetControl('GS'));
  TOBFiles := TOB.create ('LES FICHIERS',nil,-1);
  PrepareLibelle;
  GetFiles;
  PrepareLaGrid;
  AfficheLaGrid;
  SetEvents(True);
end ;

procedure TOF_BTASSOCIEDOCS.OnClose ;
begin
  TOBFiles.free;
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTASSOCIEDOCS.GetFiles;
var QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM BFILES WHERE BF0_CODE="'+LaTOB.GetString('CODE')+'"',True,-1,'',true);
  if not QQ.eof then
  begin
    TOBFiles.LoadDetailDB ('BFILES','','',QQ,false);
  end;
  Ferme(QQ);
end;

procedure TOF_BTASSOCIEDOCS.PrepareLaGrid;
begin
  GS.VidePile(false);
  //
  GS.ColCount := 2;
  GS.Cells[0,0] := ' ';
  GS.ColWidths[0] := 10;
  GS.ColAligns[0] := taCenter;
  //
  GS.Cells[1,0] := 'Définition';
  GS.ColWidths[1] := 90;
  GS.ColLengths[1] := GS.Canvas.TextWidth('W') ;
  GS.ColAligns[1] := taLeftJustify;
end;

procedure TOF_BTASSOCIEDOCS.AfficheLaGrid;
var II : Integer;
begin
  GS.RowCount := TOBFiles.detail.count +1; if GS.RowCount = 1 then GS.RowCount := 2;
  GS.FixedRows := 1;
  for II := 0 to TOBFiles.Detail.count - 1 do
  begin
    TOBFiles.detail[II].PutLigneGrid(GS,II+1,false,False,TheChamps);
  end;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GS);
end;


procedure TOF_BTASSOCIEDOCS.BVoirDocClick(Sender: TObject);
var FileName : string;
begin
  if GS.row > TOBFiles.Detail.Count then Exit;
  FileName := IncludeTrailingBackslash (EmplacementStockage)+TOBFiles.detail[GS.row-1].GetString('BF0_FILENAME');
  OuvreDocument (FileName);
end;


procedure TOF_BTASSOCIEDOCS.SetEvents(State: boolean);
begin
  if State then
  begin
    GS.OnEnter := GSEnter;
    TToolbarButton97 (GetControl('BInsert')).onclick := BInsertClick;
    TToolbarButton97 (GetControl('BDelete')).onclick := BDeleteClick;
    TToolbarButton97 (GetControl('BVOIR')).onclick := BVoirDocClick;
  end else
  begin
    GS.OnEnter := nil;
    TToolbarButton97 (GetControl('BInsert')).onclick := nil;
    TToolbarButton97 (GetControl('BDelete')).onclick := nil;
    TToolbarButton97 (GetControl('BVOIR')).onclick := nil;
  end;
end;

procedure TOF_BTASSOCIEDOCS.GSEnter(Sender: TObject);
begin
  GS.Row := 1;
end;

procedure TOF_BTASSOCIEDOCS.PrepareLibelle;
begin
  SetControlCaption('HDESCRIPTIF',LibelleAction(LaTOB));
end;

procedure TOF_BTASSOCIEDOCS.BDeleteClick(Sender: TObject);
begin
  if PGIAsk('Etes-vous sur de vous supprimer cette association')<> mryes then exit;
  if GS.row > TOBFiles.Detail.count then Exit;
  TOBFiles.Detail[Gs.Row-1].free;
  GS.DeleteRow(GS.Row); 
end;

procedure TOF_BTASSOCIEDOCS.BInsertClick(Sender: TObject);
var OneTOB : TOB;
    TTOB : TOB;
begin
  OneTOB := TOB.Create('UNE ASSO',nil,-1);
  OneTOB.AddChampSupValeur('MODIF','-');
  OneTOB.AddChampSupValeur('CODE',LaTOB.GetString('CODE'));
  TheTOB := OneTOB;
  AGLLanceFiche('BTP','BTFILESREF','','','ACTION=MODIFICATION');
  TheTOB := nil;
  if OneTOB.GetString('MODIF')='X' then
  begin
    TTOB := OneTOB.detail[0];
    TTOB.ChangeParent(TOBFiles,-1);
    TTOB.SetString('BF0_CODE',LaTOB.GetString('CODE'));
    AfficheLaGrid;
  end;
  OneTOB.free;
end;

Initialization
  registerclasses ( [ TOF_BTASSOCIEDOCS ] ) ;
end.

