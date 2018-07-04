{$IFDEF WIN32}
{$I-,R-,A-,H-}
{$ELSE}
{$I-,R-,A-}
{$ENDIF}
UNIT Gaccess ;

interface

USES DIALOGS,SYSUTILS,Classes,GDeclaDf
{$IFDEF WIN32}
,ENT1,DB,{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},HEnt1
{$ELSE}
,ExpUtil
{$ENDIF}
;

Type TGuide = RECORD
              StatusLG    : LongInt ;
              CodeLG      : String8 ;
              NoLG        : SmallInt ;
              Cpt1LG      : String13 ;
              Cpt2LG      : String13 ;
              RefLG       : String13 ;
              LibLG       : String50 ;
              DebitLg     : String40 ;
              CreditLg    : String40 ;
              END ;


PROCEDURE OUVRENUM  ;
PROCEDURE FERMENUM  ;
PROCEDURE OUVRECHOIXCOD  ;
PROCEDURE FERMECHOIXCOD  ;
PROCEDURE OUVREDEV ;
PROCEDURE FERMEDEV ;
PROCEDURE OUVRE_REGLE ;
PROCEDURE FERME_REGLE ;
PROCEDURE OUVREBILAN ( TT : Byte ) ;
PROCEDURE FERMEBILAN ( TT : Byte ) ;
FUNCTION CRECLEGUIDE ( Code : String ; No : SmallInt ) : String40 ;
PROCEDURE OUVREGUIDE ( TT : Byte ) ;
PROCEDURE FERMEGUIDE ( TT : Byte ) ;
PROCEDURE CHARGECALCMODE ( Num : Integer ; Var CalcMode : CalcModePaie ) ;
FUNCTION  GETETAT (EtatFiche : Byte ; Etat : Byte) : Boolean ;
PROCEDURE INITNEWSAA ( Sorte : smallint ; Var Fiche : Enregcpta ) ;
{$IFDEF WIN32}
PROCEDURE LITBLOC ( PosBloc : Longint ; Ch : Char ; Compta : TSaaProg ; T : TField) ;
{$ENDIF}
Function  Int2Date ( i : word ) : TDateTime;
PROCEDURE OUVRE_BLOC ( ch : char ; Compta : TSaaProg) ;
PROCEDURE FERME_BLOC ;
FUNCTION  CRECLE4 ( NoIndex : smallint ; TypeE : Byte ; Journal : String3 ;
                    DateE : Word ; NoPiece : Longint ; NoOrdre : smallint ;
                    CpteGene,CpteAux : String20 ; Simul : Byte ) : String40 ;
FUNCTION  CRECLE8 ( NoIndex : smallint ; TypeE : Byte ;
                    DateE : Word ; NoPiece : Longint ; NoOrdre : smallint ;
                    NoLigne : smallint ; Section,CpteGene : String20 ; Simul : Byte ;
                    CodeJournal : String3) : String40 ;
FUNCTION  CRECLE9 ( NoIndex : smallint ; TypeH : Byte ; NoPiece : Longint ;
                   NoEcrit : smallint ; NoEche,Nature : Byte ;
                   CpteGene,CpteAux : String20 ; DateH : Word ;
                   EtatH : Byte ; CodeLettreH : Longint ; DateP : Word ;
                   SimulH : Byte ) : String40 ;

FUNCTION  CRECLE13 ( Sorte : byte ; PosFiche : Longint ; NoOrdre : smallint) : String40 ;
FUNCTION  CRECLE14 ( PosGen : Longint ; NoOrdre : SmallInt ) : String40 ;
FUNCTION  CRECLE15 ( TypeVentil : Byte ; PosFiche : Longint ; NoOrdre : SmallInt ) : String40 ;
FUNCTION  CRECLE16 ( NoIndex : smallint ; NoTable : Byte ; CodeTranche : String17 ;
                    Intit : String35 ; Jal : String3 ; Typ : Byte ) : String40 ;
FUNCTION CRECLEBILAN ( Sorte : Byte ; Code : String4 ) : String40 ;
PROCEDURE OUVRE ( Sorte,Typ : smallint ; Reseau : Boolean ; Compta : TSaaProg ) ;
PROCEDURE FERME (sorte,typ : smallint ; reseau : boolean) ;
PROCEDURE ADDKEY(sorte : smallint ;
                 var IdxF       : IndexFile;
                 var DataRecNum : LongInt ;
                 var ProcKey                );
procedure AddRec(sorte : smallint ;
                 var DatF  : DataFile;
                 var R     : LongInt ;
                 var Buffer           );
PROCEDURE ClearKey(var IdxF : IndexFile);
PROCEDURE CloseFile(var DatF : DataFile);
PROCEDURE CloseIndex(var IdxF : IndexFile);
PROCEDURE DeleteKey(sorte : smallint ;
                    var IdxF       : IndexFile;
                    var DataRecNum : LongInt;
                    var ProcKey              );
PROCEDURE DeleteRec(sorte : smallint ;
                    var DatF : DataFile;
                        R    : LongInt);
PROCEDURE EraseFile(var DatF : DataFile);
PROCEDURE EraseIndex(var IdxF : IndexFile);
PROCEDURE FlushFile(var datF : DataFile);
PROCEDURE FlushIndex(var IdxF : IndexFile);
FUNCTION  FileLen(var DatF : DataFile) : LongInt;
PROCEDURE FindKey(var IdxF       : IndexFile;
                  var DataRecNum : LongInt;
                  var ProcKey                );
PROCEDURE GetRec(var DatF   : DataFile;
                     R      : LongInt;
                 var Buffer           );
PROCEDURE MakeFile(var DatF   : DataFile;
                       FName  : FileName;
                       RecLen : word);
PROCEDURE MakeIndex(var IdxF   : IndexFile;
                        FName  : FileName;
                        KeyLen,
                        S      : byte);
PROCEDURE NextKey(var IdxF       : IndexFile;
                  var DataRecNum : LongInt;
                  var ProcKey                );
PROCEDURE OpenFile(var DatF   : DataFile;
                       FName  : FileName;
                       RecLen : word);
PROCEDURE OpenIndex(var IdxF   : IndexFile;
                        FName  : FileName;
                        KeyLen,
                        S      : byte);
PROCEDURE PrevKey(var IdxF       : IndexFile;
                  var DataRecNum : LongInt;
                  var ProcKey                );
PROCEDURE PutRec(var DatF   : DataFile;
                     R      : LongInt;
                 var Buffer           );
PROCEDURE SearchKey(var IdxF : IndexFile;
                    var DataRecNum : LongInt;
                    var ProcKey);
PROCEDURE SearchPrevKey(var IdxF : IndexFile;
                        var DataRecNum : LongInt;
                        var ProcKey);
FUNCTION  UsedRecs(var DatF : DataFile) : LongInt;
PROCEDURE TACrash(FileNm : FileName;
                  R : LongInt);


IMPLEMENTATION

var
  TaRecBuf  : TaRecordBufPtr;
  TaPageStk : TaPageStackPtr;
  TaPgMap   : TaPageMapPtr;

{===============================================================}
PROCEDURE TACrash(FileNm : FileName;
                  R : LongInt);
BEGIN
(*
ShowMessage(IntToStr(VSAA^.TaStatus)) ;
*)
(*
  case TAstatus of
      1..6: TurboError := mess(51+TaStatus) ;
      8: TurboError := mess(58) ;
     12: TurboError := mess(59) ;
     15..17: TurboError := mess(45+TaStatus) ;
    100..106: TurboError := mess(TaStatus-37) ;
    200..204: TurboError := mess(TAStatus-130) ;
    RecTooLarge : TurboError := mess(75) ;
    RecTooSmall : TurboError := mess(76) ;
    KeyTooLarge : TurboError := mess(77) ;
    RecSizeMismatch : TurboError := mess(78) ;
    KeySizeMismatch : TurboError := mess(79) ;
    MemOverflow : TurboError := mess(80) ;
    else
       TurboError := '';

    END;
  ModeMenu:=0 ;
{$IFDEF WINDOWS}
{$ELSE}
  prepfen(1,21,79,24,lightred) ;  rep(1) ;
  color(yellow,black) ;
  gotoxy(20,22) ; write('TAErreur nø ',tastatus,' --> ',turboerror) ;
  gotoxy(20,23) ; Writeln('File : ', FileNm,'  Record nø ', R);
  color(white,black) ; gotoxy(1,24) ;
{$ENDIF}
  Halt(0);
*)
END; { TACrash }

{===============================================================}
PROCEDURE TAIOCheck(var DatF : DataFile;
                    R : LongInt);
var
  i : byte;
  FileNm : FileName;
BEGIN
  if VSAA^.TAStatus <> 0 then
    with DatF do
    BEGIN
      i := 0;
      FileNm := '';
      with TFileRec(F) do
        while (Ord(Name[i]) <> 0) and (i <= 80) do
        BEGIN
          FileNm := FileNm + Name[i];
          i := Succ(i);
        END;
        TACrash(FileNm, R);
     END;
END; { TAIOCheck }

{===============================================================}
PROCEDURE GetRec(var DatF   : DataFile;
                     R      : LongInt;
                 var Buffer           );
var
  BlocksRead : Integer ;
BEGIN
  Seek(DatF.F,R);
  VSAA^.TAStatus := IOresult;
  TaIOcheck(DatF,R);
  BlockRead(DatF.F,Buffer, 1, BlocksRead);
  if BlocksRead = 0 then VSAA^.TAStatus := 100;
  TaIOcheck(DatF, R);
END; { GetRec }

{===============================================================}
PROCEDURE PutRec(var DatF   : DataFile;
                     R      : LongInt;
                 var Buffer           );
var
  BlocksWritten : Integer ;

BEGIN

{ M9 }
{
If ADM AND FICADM(Enregistrement(Buffer).Sorte) Then Exit ;
}

Seek(DatF.F,R);
VSAA^.TAStatus := IOresult;
TaIOcheck(DatF,R);
BlockWrite(DatF.F,Buffer,1, BlocksWritten);
if BlocksWritten=0 then VSAA^.TAStatus := 101 ;
TaIOcheck(DatF,R);
END; { PutRec }

{===============================================================}
PROCEDURE GetZero(var DatF   : DataFile) ;

BEGIN
if VSAA^.WarnIndex then exit ;
    FlushFile(Datf) ;
    TaIOcheck(DatF,0);
    GetRec(DatF,0,TaRecBuf^);
    Move(TaRecBuf^,DatF.FirstFree,fileheadersize);
    DatF.NumRec := FileSize(DatF.F);
    if ((datf.firstfree<=0) or (datf.firstfree>=datf.numrec)) then
       BEGIN
       datf.firstfree:=-1 ;
       datf.numberfree:=0 ;
       END ;
END;

{===============================================================}
PROCEDURE PutZero(var DatF : DataFile);
BEGIN
Exit ; { GG WIN }
if VSAA^.WarnIndex then exit ;
if ((datf.firstfree<=0) or (datf.firstfree>=datf.numrec)) then
     BEGIN datf.numberFree:=0; datf.FirstFree:=-1; END ;
  Move(DatF.FirstFree,TaRecBuf^,FileHeaderSize);
  PutRec(DatF,0,TaRecBuf^);
  TaIOcheck(DatF,0);
  FlushFile(Datf) ;
END;


{===============================================================}
PROCEDURE WriteHeader(var DatF   : DataFile; RecLen : word);

BEGIN
  with DatF do
  BEGIN
    FillChar(TaRecBuf^, SizeOf(TaRecBuf^), 0);
    FirstFree := -1;
    ItemSize := RecLen;
    Move(DatF.FirstFree,TaRecBuf^,FileHeaderSize);
    PutRec(DatF,0,TaRecBuf^);
    TaIOCheck(DatF, 0);
    DatF.NumRec := 1;
    VSAA^.OK := true;
  END;
END; { WriteHeader }

{===============================================================}
PROCEDURE MakeFile(var DatF   : DataFile;
                       FName  : FileName;
                       RecLen : word);
BEGIN
  VSAA^.TAStatus := 0;
  FillChar(DatF, SizeOf(DatF), 0);
  Assign(DatF.F, FName);
  Rewrite(DatF.F,RecLen);
  VSAA^.Ok := IOResult = 0;
  if VSAA^.Ok then
    BEGIN
    if RecLen > MaxDataRecSize then VSAA^.TAStatus := RecTooLarge;
    if RecLen < MinDataRecSize then VSAA^.TAStatus := RecTooSmall;
    TAIOCheck(DatF, 0);
    WriteHeader(DatF, RecLen);
  END;
END; { MakeFile }

{===============================================================}
PROCEDURE ReadHeader(var DatF   : DataFile);

BEGIN
  with DatF do
  BEGIN
    GetRec(DatF,0,TaRecBuf^);
    Move(TaRecBuf^,DatF.FirstFree,FileHeaderSize);
    DatF.NumRec := FileSize(DatF.F);
  END;
END; { ReadHeader }

{===============================================================}
PROCEDURE OpenFile(var DatF   : DataFile;
                       FName  : FileName;
                       RecLen : word);
BEGIN
  FillChar(DatF, SizeOf(DatF), 0);
  Assign(DatF.F,FName);
  Reset(DatF.F,RecLen);
  VSAA^.TAStatus := IOResult;
  VSAA^.Ok := VSAA^.TAStatus = 0;
  if VSAA^.Ok then
    BEGIN
    if RecLen > MaxDataRecSize then VSAA^.TAStatus := RecTooLarge;
    if RecLen < MinDataRecSize then VSAA^.TAStatus := RecTooSmall;
    TAIOCheck(DatF, 0);
    ReadHeader(DatF);
    if datf.firstfree<=0 then BEGIN datf.numberFree:=0; datf.FirstFree:=-1; END ;
    END;
END; { OpenFile }

{===============================================================}
PROCEDURE PutFileHeader(var DatF : DataFile);

BEGIN
  Exit ; { GG WIN }
  FillChar(TaRecBuf^, SizeOf(TaRecBuf^), 0);
  Move(DatF.FirstFree,TaRecBuf^,FileHeaderSize);
  PutRec(DatF,0,TaRecBuf^);
END; { PutFileHeader }

{===============================================================}
PROCEDURE CloseFile(var DatF : DataFile);
Label 1 ;
BEGIN
  if datf.firstFree<=0 then
     BEGIN
     datf.numberFree:=0;
     datf.FirstFree:=-1;
     END ;
(*???  if NbPosteReseau<2 then PutFileHeader(DatF);*)
1:
  Close(DatF.F);
  VSAA^.TAStatus := IOresult;
  TaIOcheck(DatF,0);
END; { CloseFile }

{===============================================================}
PROCEDURE FlushFile(var DatF : DataFile);

BEGIN
(*
  if NbPosteReseau<2 then   PutFileHeader(DatF);
  FlushDOSFile(DatF.F);
  TAIOCheck(DatF, 0);
*)
END; { FlushFile }

{===============================================================}
PROCEDURE NewRec(var DatF  : DataFile;
                 var R     : LongInt  );
BEGIN
  if DatF.FirstFree<= -1 then
  BEGIN
    R := DatF.NumRec;
    DatF.NumRec := Succ(DatF.NumRec) ;
    datf.numberFree:=0;
    datf.FirstFree:=-1;
  END
  else
  BEGIN
    R := DatF.FirstFree;
    GetRec(DatF,R,TaRecBuf^);
    DatF.FirstFree := TaRecBuf^.I;
    DatF.NumberFree := pred(DatF.NumberFree);
  END;
(*???if NbPosteReseau>1 then putzero(datf) ;*)
END; { NewRec }

{===============================================================}
PROCEDURE AddRec(sorte : smallint ;
                 var DatF  : DataFile;
                 var R     : LongInt;
                 var Buffer           );

Var Status : Longint ;
    Reseau : Byte ;
    Tab : Array [1..6] of Byte ;

BEGIN
(*???if NbPosteReseau>1 then getzero(datf) ;*)
NewRec(DatF,R);
if Sorte>=0 then
   BEGIN
   Status:=0 ; Move(Status,Tab[1],4) ;
   Reseau:=0 ; Move(reseau,Tab[5],1) ;
   Move(Sorte,Tab[6],1) ; Move(Tab[1],Buffer,6) ;
   END ;
PutRec(DatF,R,Buffer);
END; { AddRec }

{===============================================================}
PROCEDURE DeleteRec(sorte : smallint ;
                    var DatF : DataFile;
                        R    : LongInt);
BEGIN
(*???if NbPosteReseau>1 then getzero(datf) ;*)
TaRecBuf^.I := DatF.FirstFree;
PutRec(DatF,R,TaRecBuf^);
if R<=0 then R:=-1 ;
DatF.FirstFree := R;
DatF.NumberFree := succ(DatF.NumberFree);
(*???if NbPosteReseau>1 then putzero(datf) ;*)
END; { DeleteRec }

{===============================================================}
FUNCTION FileLen(var DatF : DataFile) : LongInt;

BEGIN
  FileLen := DatF.NumRec;
END; { FileLen }

{===============================================================}
FUNCTION UsedRecs(var DatF : DataFile) : LongInt;

BEGIN
  UsedRecs := DatF.NumRec - DatF.NumberFree - 1;
END; { UsedRecs }

Type
  PageBlock = array[0..32767] of Byte;

{===============================================================}
PROCEDURE TaPack(var Page : TaPage;KeyL : byte);
Var I : word ;
    P : ^PageBlock;

BEGIN
  P := @Page;
  if KeyL <> MaxKeyLen then
    for I := 1 to PageSize do
      Move(Page.ItemArray[I], P^[(I - 1) * (KeyL + ItemOverhead)
                                 + PageOverhead],KeyL + ItemOverhead);
END; { TaPack }

{===============================================================}
PROCEDURE TaUnpack(var Page : TaPage;
                       KeyL : byte);
Var I : word;
    P : ^PageBlock;

BEGIN
  P := @Page;
  if KeyL <> MaxKeyLen then
    for I := PageSize downto 1 do
      Move(P^[(I - 1) * (KeyL + ItemOverhead) + PageOverhead],
           Page.ItemArray[I],KeyL + ItemOverhead);
END; { TaUnpack }

{===============================================================}
PROCEDURE OpenIndex(var IdxF   : IndexFile;
                        FName  : FileName;
                        KeyLen,
                        S      : byte);
Var K : word;

BEGIN
  K := (KeyLen + ItemOverhead) * PageSize + PageOverhead;
  FillChar(IdxF, SizeOf(IdxF), 0); IdxF.DelAdd:=FALSE ;
  with IdxF, DataF do
    BEGIN
    Assign(F, FName);
    Reset(F, K);
    END;
  VSAA^.TAStatus := IOResult;
  VSAA^.Ok := VSAA^.TAStatus = 0;
  if VSAA^.Ok then
    BEGIN
    if KeyLen > MaxKeyLen then
      BEGIN
      VSAA^.TAStatus := KeyTooLarge;
      TAIOCheck(IdxF.DataF, 0);
      END;
    ReadHeader(IdxF.DataF);
    if K <> IdxF.DataF.ItemSize then
      BEGIN
      VSAA^.TAStatus := KeySizeMismatch;
      TAIOCheck(IdxF.DataF, 0);
      END;
    IdxF.AllowDuplKeys := S <> NoDuplicates;
    IdxF.KeyL := KeyLen;
    IdxF.RR := IdxF.DataF.Int1;
    IdxF.PP := 0;
    END ;
END; { OpenIndex }

{===============================================================}
PROCEDURE StoreIndexHeader(var IdxF : IndexFile; OkMake:boolean);

Var I : smallint;
    gg : boolean ;

BEGIN
gg:=FALSE ;
Exit ; { GG win }
for I := 1 to PageStackSize do
  with TaPageStk^[I] do
    if IndexFPtr = @IdxF then
      BEGIN
      IndexFPtr := nil;
      if Updated then
        BEGIN
        TaPack(Page,IdxF.KeyL);
        PutRec(IdxF.DataF,PageRef,Page);
        Updated := false; gg:=TRUE ;
        END;
      END;
  if (gg or okmake or IdxF.DelAdd) then
     BEGIN
     IdxF.DataF.Int1 := IdxF.RR;
     PutFileHeader(IdxF.DataF);
     END ;
END; { StoreIndexHeader }

{===============================================================}
PROCEDURE CloseIndex(var IdxF : IndexFile);

BEGIN
  StoreIndexHeader(IdxF,FALSE);
  VSAA^.WarnIndex:=TRUE ;
  CloseFile(IdxF.DataF);
  VSAA^.WarnIndex:=FALSE ;
END; { CloseIndex }

{===============================================================}
PROCEDURE MakeIndex(var IdxF   : IndexFile;
                        FName  : FileName;
                        KeyLen,
                        S      : byte);
Var K : word;

BEGIN
  FillChar(IdxF, SizeOf(IdxF), 0);
  K := (KeyLen + ItemOverhead) * PageSize + PageOverhead;
  with IdxF,IdxF.DataF do
    BEGIN
    Assign(F, FName);
    Rewrite(F, K);
    VSAA^.TAStatus := IOResult;
    if KeyLen > MaxKeyLen then VSAA^.TAStatus := KeyTooLarge;
    TAIOcheck(DataF, 0);
    WriteHeader(IdxF.DataF, K);
    IdxF.AllowDuplKeys := S <> NoDuplicates;
    IdxF.KeyL := KeyLen;
    END;
StoreIndexHeader(IdxF,TRUE);
END; { MakeIndex }

{===============================================================}
PROCEDURE FlushIndex(var IdxF       : IndexFile);

BEGIN
  StoreIndexHeader(IdxF,TRUE);
  FlushFile(IdxF.DataF);
END; { FlushIndex }

{===============================================================}
PROCEDURE EraseFile(var DatF : DataFile);

BEGIN
  CloseFile(DatF);
  Erase(DatF.F);
END; { EraseFile }

{===============================================================}
PROCEDURE EraseIndex(var IdxF : IndexFile);

BEGIN
  CloseIndex(IdxF);
  Erase(IdxF.DataF.F);
END; { EraseIndex }

{===============================================================}
PROCEDURE TaLast(I : LongInt);

Var J,K : smallint;

BEGIN
  J := 1;
  while (TaPgMap^[J] <> I) and (J < PageStackSize) do
    J := succ(J);
  for K := J to pred(PageStackSize)do
    TaPgMap^[K] := TaPgMap^[succ(K)];
  TaPgMap^[PageStackSize] := I;
END; { TaLast }

{===============================================================}
PROCEDURE TaGetPage(var IdxF  : IndexFile;
                        R     : LongInt ;
                    var PgPtr : TaPagePtr);
Var I     : smallint;
    Found : Boolean;

BEGIN
  I := 0;
  repeat
    I := succ(I);
    with TaPageStk^[I] do
      Found := (IndexFPtr = @IdxF) and (PageRef = R);
  until Found or (I = PageStackSize);

  if not Found then
  BEGIN
    I := TaPgMap^[1];
    with TaPageStk^[I] do
    BEGIN
      if Updated then
      BEGIN
        TaPack(Page,IndexFPtr^.KeyL);
        PutRec(IndexFPtr^.DataF,PageRef,Page);
      END;
      GetRec(IdxF.DataF,R,Page);
      TaUnpack(Page,IdxF.KeyL);
      IndexFPtr := @IdxF;
      PageRef := R;
      Updated := false;
    END;
  END;
  TaLast(I);
  PgPtr := @TaPageStk^[I];
END; { TaGetPage }

{===============================================================}
PROCEDURE TaNewPage(var IdxF  : IndexFile;
                    var R     : LongInt;
                    var PgPtr : TaPagePtr);
Var I : smallint;

BEGIN
  I := TaPgMap^[1];
  with TaPageStk^[I] do
  BEGIN
    if Updated then
    BEGIN
      TaPack(Page,IndexFPtr^.KeyL);
      PutRec(IndexFPtr^.DataF,PageRef,Page);
    END;
    VSAA^.WarnIndex:=TRUE ;
    NewRec(IdxF.DataF,R);
    VSAA^.WarnIndex:=FALSE ;
    IndexFPtr := @IdxF;
    PageRef := R;
    Updated := false;
  END;
  TaLast(I);
  PgPtr := @TaPageStk^[I];
END; { TANewPage }

{===============================================================}
PROCEDURE TaUpdatePage(PgPtr : TaPagePtr);

BEGIN
  TAStackRecPtr(PgPtr)^.Updated := true;
END; { TAUpdatePage }

{===============================================================}
PROCEDURE TaReturnPage(var PgPtr : TaPagePtr);

BEGIN
  with TaStackRecPtr(PgPtr)^ do
  BEGIN
    VSAA^.Warnindex:=TRUE ;
    DeleteRec(-1,IndexFPtr^.DataF,PageRef);
    VSAA^.Warnindex:=FALSE ;
    IndexFPtr := nil;
    Updated := FALSE ;
  END;
END; { TAReturnPage }

{===============================================================}
PROCEDURE TaXKey(var K;
                 KeyL : byte);
BEGIN
  if Length(TaKeyStr(K)) > KeyL then
    TaKeyStr(K)[0] := Chr(KeyL);
END; { TaXKey }

{===============================================================}
FUNCTION TaCompKeys(var K1,
                        K2;
                        DR1,
                        DR2 : LongInt;
                        Dup : Boolean ) : Longint ;
BEGIN
  if TaKeyStr(K1) = TaKeyStr(K2) then
    if Dup then TaCompKeys := DR1 - DR2
           else TaCompKeys := 0
  else
    if TaKeyStr(K1) > TaKeyStr(K2) then TaCompKeys := 1
                                   else TaCompKeys :=  - 1;
END;

{===============================================================}
PROCEDURE ClearKey(var IdxF : IndexFile);

BEGIN
IdxF.PP:=0 ;
END;

{===============================================================}
PROCEDURE NextKey(var IdxF       : IndexFile;
                  var DataRecNum : LongInt;
                  var ProcKey                );
Var R      : LongInt;
    PagPtr : TaPagePtr;

BEGIN
  with IdxF do
  BEGIN
    if PP = 0 then R := RR
    else
    with Path[PP] do
    BEGIN
      TaGetPage(IdxF,PageRef,PagPtr);
      R := PagPtr^.ItemArray[ItemArrIndex].PageRef;
    END;
    while R <> 0 do
    BEGIN
      PP := PP + 1;
      with Path[PP] do
      BEGIN
        PageRef := R;
        ItemArrIndex := 0;
      END;
      TaGetPage(IdxF,R,PagPtr);
      R := PagPtr^.BckwPageRef;
    END;
    if PP <> 0 then
    BEGIN
      while (PP > 1) and
            (Path[PP].ItemArrIndex = PagPtr^.ItemsOnPage) do
      BEGIN
        PP := PP - 1;
        TaGetPage(IdxF,Path[PP].PageRef,PagPtr);
      END;
      if Path[PP].ItemArrIndex < PagPtr^.ItemsOnPage then
        with Path[PP] do
        BEGIN
          ItemArrIndex := ItemArrIndex + 1;
          with PagPtr^.ItemArray[ItemArrIndex] do
          BEGIN
            TaKeyStr(ProcKey) := Key;
            DataRecNum := DataRef;
          END;
        END
      else PP := 0;
    END;
    VSAA^.OK := PP <> 0;
  END;
END; { NextKey }

{===============================================================}
PROCEDURE PrevKey(var IdxF       : IndexFile;
                  var DataRecNum : LongInt;
                  var ProcKey                );
Var R      : LongInt;
    PagPtr : TaPagePtr;

BEGIN
  with IdxF do
  BEGIN
    if PP = 0 then
      R := RR
    else
      with Path[PP] do
      BEGIN
        TaGetPage(IdxF,PageRef,PagPtr);
        ItemArrIndex := ItemArrIndex - 1;
        if ItemArrIndex = 0 then
          R := PagPtr^.BckwPageRef
        else R := PagPtr^.ItemArray[ItemArrIndex].PageRef;
      END;
    while R <> 0 do
    BEGIN
      TaGetPage(IdxF,R,PagPtr);
      PP := PP + 1;
      with Path[PP] do
      BEGIN
        PageRef := R;
        ItemArrIndex := PagPtr^.ItemsOnPage;
      END;
      with PagPtr^ do
        R := ItemArray[ItemsOnPage].PageRef;
    END;
    if PP <> 0 then
    BEGIN
      while (PP > 1) and (Path[PP].ItemArrIndex = 0) do
      BEGIN
        PP := PP - 1;
        TaGetPage(IdxF,Path[PP].PageRef,PagPtr);
      END;
      if Path[PP].ItemArrIndex > 0 then
        with PagPtr^.ItemArray[Path[PP].ItemArrIndex] do
        BEGIN
          TaKeyStr(ProcKey) := Key;
          DataRecNum := DataRef;
        END
      else PP := 0;
    END;
    VSAA^.OK := PP <> 0;
  END;
END; { PrevKey }

{===============================================================}
PROCEDURE TaFindKey(var IdxF       : IndexFile;
                    var DataRecNum : LongInt;
                    var ProcKey                );

Var   C,PrPgRef : LongInt;
      K,L,R     : smallint;
      PagPtr    : TaPagePtr;

BEGIN
  with IdxF do
  BEGIN
    TaXKey(TaKeyStr(ProcKey),KeyL);
    VSAA^.OK := false;
    PP := 0;
    PrPgRef := RR;
    while (PrPgRef <> 0) and not VSAA^.OK do
    BEGIN
      PP := PP + 1;
      Path[PP].PageRef := PrPgRef;
      TaGetPage(IdxF,PrPgRef,PagPtr);
      with PagPtr^ do
      BEGIN
        L := 1;
        R := ItemsOnPage;
        repeat
          K := (L + R) div 2;
          C := TaCompKeys(TaKeyStr(ProcKey),
                          ItemArray[K].Key,
                          0,
                          ItemArray[K].DataRef,
                          AllowDuplKeys        );
          if C <= 0 then R := K - 1;
          if C >= 0 then L := K + 1;
        until R < L;
        if L - R > 1 then
          BEGIN
          DataRecNum := ItemArray[K].DataRef;
          R := K;
          VSAA^.OK := true;
          END;
        if R = 0 then PrPgRef := BckwPageRef
                 else PrPgRef := ItemArray[R].PageRef;
      END;
      Path[PP].ItemArrIndex := R;
    END;
    if not VSAA^.OK and (PP > 0) then
    BEGIN
      while (PP > 1) and (Path[PP].ItemArrIndex = 0) do
        PP := PP - 1;
      if Path[PP].ItemArrIndex = 0 then
        PP := 0;
    END;
  END;
END; { TAFindKey }

{===============================================================}
PROCEDURE FINDKEY ( Var IdxF : IndexFile ; Var DataRecNum : Longint ;
                    Var ProcKey ) ;
Var TempKey : TaKeyStr;
BEGIN
TaFindKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
if Not VSAA^.Ok and IdxF.AllowDuplKeys then
   BEGIN
   TempKey:=TaKeyStr(ProcKey) ;
   NextKey(IdxF,DataRecNum,TaKeyStr(ProcKey)) ;
   VSAA^.Ok:=VSAA^.Ok and (TaKeyStr(ProcKey)=TempKey) ;
   END ;
END;

{===============================================================}
PROCEDURE SearchKey(var IdxF : IndexFile;
                    var DataRecNum : LongInt;
                    var ProcKey);
BEGIN
  TaFindKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
  if not VSAA^.OK then NextKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
END; { SearchKey }

{===============================================================}
PROCEDURE SearchPrevKey(var IdxF : IndexFile;
                        var DataRecNum : LongInt;
                        var ProcKey);
Var Proc2 : String ;
BEGIN
TaFindKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
if not VSAA^.OK then
   BEGIN
   if IdxF.AllowDuplKeys then
      BEGIN
      Proc2:=TaKeyStr(ProcKey) ;
      Repeat
       NextKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
      Until ((not VSAA^.ok) or (TaKeyStr(ProcKey)<>Proc2)) ;
      PrevKey(IdxF,DataRecNum,TaKeyStr(ProcKey)) ;
      END else
      BEGIN
      PrevKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
      NextKey(IdxF,DataRecNum,TaKeyStr(ProcKey));
      END ;
   END ;
END; { SearchPrevKey }

{===============================================================}
PROCEDURE AddKey(sorte : smallint ;
                 var IdxF       : IndexFile;
                 var DataRecNum : LongInt;
                 var ProcKey                );
Var PrPgRef1,C,
    PrPgRef2  : LongInt;
    K,L   : smallint;
    PassUp    : Boolean;
    PagePtr1,
    PagePtr2  : TaPagePtr;
    ProcItem1,
    ProcItem2 : TaItem;

{111111111111111111111111111111111111111111111}
PROCEDURE Search(PrPgRef1 : LongInt);
Var R : smallint;

{222222222222222222222222222222222222222222222}
PROCEDURE Insert;
Var i : smallint;

BEGIN
  TaGetPage(IdxF,PrPgRef1,PagePtr1);
  with PagePtr1^ do
  BEGIN
    if ItemsOnPage < PageSize then
    BEGIN
      ItemsOnPage := ItemsOnPage + 1;
      for I := ItemsOnPage downto R + 2 do
        ItemArray[I] := ItemArray[I - 1];
      ItemArray[R + 1] := ProcItem1;
      PassUp := false;
    END
    else
    BEGIN
      TaNewPage(IdxF,PrPgRef2,PagePtr2);
      if R <= Order then
      BEGIN
        if R = Order then
          ProcItem2 := ProcItem1
        else
        BEGIN
          ProcItem2 := ItemArray[Order];
          for I := Order downto R + 2 do
            ItemArray[I] := ItemArray[I - 1];
          ItemArray[R + 1] := ProcItem1;
        END;
        for I := 1 to Order do
          PagePtr2^.ItemArray[I] := ItemArray[I + Order];
      END
      else
      BEGIN
        R := R - Order;
        ProcItem2 := ItemArray[Order + 1];
        for I := 1 to R - 1 do
          PagePtr2^.ItemArray[I] := ItemArray[I + Order + 1];
        PagePtr2^.ItemArray[R] := ProcItem1;
        for I := R + 1 to Order do
          PagePtr2^.ItemArray[I] := ItemArray[I + Order];
      END;
      ItemsOnPage := Order;
      PagePtr2^.ItemsOnPage := Order;
      PagePtr2^.BckwPageRef := ProcItem2.PageRef;
      ProcItem2.PageRef := PrPgRef2;
      ProcItem1 := ProcItem2;
      TaUpdatePage(PagePtr2);
    END;
  END;
  TaUpdatePage(PagePtr1);
END; { Insert }

BEGIN { Search }
  if PrPgRef1 = 0 then
  BEGIN
    PassUp := true;
    with ProcItem1 do
    BEGIN
      Key := TaKeyStr(ProcKey);
      DataRef := DataRecNum;
      PageRef := 0;
    END;
  END
  else
  BEGIN
    TaGetPage(IdxF,PrPgRef1,PagePtr1);
    with PagePtr1^ do
    BEGIN
      L := 1;
      R := ItemsOnPage;
      repeat
        K := (L + R) div 2;
          C := TaCompKeys(TaKeyStr(ProcKey),
                        ItemArray[K].Key,
                        DataRecNum,
                        ItemArray[K].DataRef,
                        IdxF.AllowDuplKeys   );
        if C <= 0 then
          R := K - 1;
        if C >= 0 then
          L := K + 1;
      until R < L;
      if L - R > 1 then
      BEGIN
        VSAA^.OK := false;
        PassUp := false;
      END
      else
      BEGIN
        if R = 0 then
          Search(BckwPageRef)
        else Search(ItemArray[R].PageRef);
        if PassUp then
          Insert;
      END;
    END;
  END;
END; { Search }

BEGIN { AddKey }
IdxF.DelAdd:=TRUE ;
with IdxF do
  BEGIN
    TaXKey(TaKeyStr(ProcKey),KeyL);
    VSAA^.OK := true;
    Search(RR);
    if PassUp then
    BEGIN
      PrPgRef1 := RR;
      TaNewPage(IdxF,RR,PagePtr1);
      with PagePtr1^ do
      BEGIN
        ItemsOnPage := 1;
        BckwPageRef := PrPgRef1;
        ItemArray[1] := ProcItem1;
      END;
      TaUpdatePage(PagePtr1);
    END;
    PP := 0;
  END;
END; { AddKey }

{===============================================================}
PROCEDURE DeleteKey(sorte : smallint ;
                    var IdxF       : IndexFile;
                    var DataRecNum : LongInt;
                    var ProcKey              );
Var PageTooSmall : Boolean;
    PagPtr       : TaPagePtr;

{11111111111111111111111111111111111111}
PROCEDURE DelB(PrPgRef : LongInt);

Var C,I,K,L,R,
    XPageRef  : LongInt;
    PagPtr    : TaPagePtr;

{22222222222222222222222222222222222222}
PROCEDURE UnderFlow(PrPgRef,
                    PrPgRef2 : LongInt;
                    R        : smallint);
Var I,K,
    LItem : smallint;
    LPageRef : LongInt;
    PagPtr,
    PagePtr2,
    L        : TaPagePtr;

BEGIN
  TaGetPage(IdxF,PrPgRef,PagPtr);
  TaGetPage(IdxF,PrPgRef2,PagePtr2);
  if R < PagPtr^.ItemsOnPage then
  BEGIN
    R := R + 1;
    LPageRef := PagPtr^.ItemArray[R].PageRef;
    TaGetPage(IdxF,LPageRef,L);
    K := (L^.ItemsOnPage - Order + 1) div 2;
    PagePtr2^.ItemArray[Order] := PagPtr^.ItemArray[R];
    PagePtr2^.ItemArray[Order].PageRef := L^.BckwPageRef;
    if K > 0 then
    BEGIN
      for I := 1 to K - 1 do
        PagePtr2^.ItemArray[I + Order] := L^.ItemArray[I];
      PagPtr^.ItemArray[R] := L^.ItemArray[K];
      PagPtr^.ItemArray[R].PageRef := LPageRef;
      L^.BckwPageRef := L^.ItemArray[K].PageRef;
      L^.ItemsOnPage := L^.ItemsOnPage - K;
      for I := 1 to L^.ItemsOnPage do
        L^.ItemArray[I] := L^.ItemArray[I + K];
      PagePtr2^.ItemsOnPage := Order - 1 + K;
      PageTooSmall := false;
      TaUpdatePage(L);
    END
    else
    BEGIN
      for I := 1 to Order do
        PagePtr2^.ItemArray[I + Order] := L^.ItemArray[I];
      for I := R to PagPtr^.ItemsOnPage - 1 do
        PagPtr^.ItemArray[I] := PagPtr^.ItemArray[I + 1];
      PagePtr2^.ItemsOnPage := PageSize;
      PagPtr^.ItemsOnPage := PagPtr^.ItemsOnPage - 1;
      TaReturnPage(L);
      PageTooSmall := PagPtr^.ItemsOnPage < Order;
    END;
    TaUpdatePage(PagePtr2);
  END
  else
  BEGIN
    if R = 1 then
      LPageRef := PagPtr^.BckwPageRef
    else LPageRef := PagPtr^.ItemArray[R - 1].PageRef;
    TaGetPage(IdxF,LPageRef,L);
    LItem := L^.ItemsOnPage + 1;
    K := (LItem - Order) div 2;
    if K > 0 then
    BEGIN
      for I := Order - 1 downto 1 do
        PagePtr2^.ItemArray[I + K] := PagePtr2^.ItemArray[I];
      PagePtr2^.ItemArray[K] := PagPtr^.ItemArray[R];
      PagePtr2^.ItemArray[K].PageRef := PagePtr2^.BckwPageRef;
      LItem := LItem - K;
      for I := K - 1 downto 1 do
        PagePtr2^.ItemArray[I] := L^.ItemArray[I + LItem];
      PagePtr2^.BckwPageRef := L^.ItemArray[LItem].PageRef;
      PagPtr^.ItemArray[R] := L^.ItemArray[LItem];
      PagPtr^.ItemArray[R].PageRef := PrPgRef2;
      L^.ItemsOnPage := LItem - 1;
      PagePtr2^.ItemsOnPage := Order - 1 + K;
      PageTooSmall := false;
      TaUpdatePage(PagePtr2);
    END
    else
    BEGIN
      L^.ItemArray[LItem] := PagPtr^.ItemArray[R];
      L^.ItemArray[LItem].PageRef := PagePtr2^.BckwPageRef;
      for I := 1 to Order - 1 do
        L^.ItemArray[I + LItem] := PagePtr2^.ItemArray[I];
      L^.ItemsOnPage := PageSize;
      PagPtr^.ItemsOnPage := PagPtr^.ItemsOnPage - 1;
      TaReturnPage(PagePtr2);
      PageTooSmall := PagPtr^.ItemsOnPage < Order;
    END;
    TaUpdatePage(L);
  END;
  TaUpdatePage(PagPtr);
END; { UnderFlow }

{3333333333333333333333333333333333333333}
PROCEDURE DelA(PrPgRef2 : LongInt);

Var C,XPageRef : LongInt;
    PagePtr2 : TaPagePtr;

BEGIN
  TaGetPage(IdxF,PrPgRef2,PagePtr2);
  with PagePtr2^ do
  BEGIN
    XPageRef := ItemArray[ItemsOnPage].PageRef;
    if XPageRef <> 0 then
    BEGIN
      C := ItemsOnPage;
      DelA(XPageRef);
      if PageTooSmall then UnderFlow(PrPgRef2,XPageRef,C);
    END
    else
    BEGIN
      TaGetPage(IdxF,PrPgRef,PagPtr);
      ItemArray[ItemsOnPage].PageRef := PagPtr^.ItemArray[K].PageRef;
      PagPtr^.ItemArray[K] := ItemArray[ItemsOnPage];
      ItemsOnPage := ItemsOnPage - 1;
      PageTooSmall := ItemsOnPage < Order;
      TaUpdatePage(PagPtr);
      TaUpdatePage(PagePtr2);
    END;
  END;
END; { DelA }

BEGIN { DelB }
  if PrPgRef = 0 then
  BEGIN
    VSAA^.OK := false;
    PageTooSmall := false;
  END
  else
  BEGIN
    TaGetPage(IdxF,PrPgRef,PagPtr);
    with PagPtr^ do
    BEGIN
      L := 1;
      R := ItemsOnPage;
      repeat
        K := (L + R) div 2;
          C := TaCompKeys(TaKeyStr(ProcKey),
                        ItemArray[K].Key,
                        DataRecNum,
                        ItemArray[K].DataRef,
                        IdxF.AllowDuplKeys   );
        if C <= 0 then
          R := K - 1;
        if C >= 0 then
          L := K + 1;
      until L > R;
      if R = 0 then
        XPageRef := BckwPageRef
      else XPageRef := ItemArray[R].PageRef;
      if L - R > 1 then
      BEGIN
        DataRecNum := ItemArray[K].DataRef;
        if XPageRef = 0 then
        BEGIN
          ItemsOnPage := ItemsOnPage - 1;
          PageTooSmall := ItemsOnPage < Order;
          for I := K to ItemsOnPage do
            ItemArray[I] := ItemArray[I + 1];
          TaUpdatePage(PagPtr);
        END
        else
        BEGIN
          DelA(XPageRef);
          if PageTooSmall then
            UnderFlow(PrPgRef,XPageRef,R);
        END;
      END
      else
      BEGIN
        DelB(XPageRef);
        if PageTooSmall then
          UnderFlow(PrPgRef,XPageRef,R);
      END;
    END;
  END;
END; { DelB }

BEGIN { DeleteKey }
  IdxF.DelAdd:=TRUE ;
  with IdxF do
    BEGIN
    TaXKey(TaKeyStr(ProcKey),KeyL);
    VSAA^.OK := true;
    DelB(RR);
    if PageTooSmall then
      BEGIN
      TaGetPage(IdxF,RR,PagPtr);
      if PagPtr^.ItemsOnPage = 0 then
        BEGIN
        RR := PagPtr^.BckwPageRef;
        TaReturnPage(PagPtr);
        END;
      END;
    PP := 0;
    END;
END; { DeleteKey }

{===============================================================}
PROCEDURE InitTAccess;

{1111111111111111111111111111111}
PROCEDURE InitPages;
{ Allocate space for the page stack, the page map and the record
  buffer. }

Var i : word ;

BEGIN
(*
  if MaxAvail < (SizeOf(TaRecordBuffer) +
                 SizeOf(TaPageMap) +
                 SizeOf(TaPageStack)) then
 BEGIN
   VSAA^.TAStatus := MemOverflow;
   TACrash('', 0);
 END;
 *)
  new(TaPageStk);
  FillChar(TaPageStk^, SizeOf(TaPageStk^), 0);
  new(TaPgMap);
  for i := 1 to PageStackSize do
    TaPgMap^[i] := i;
  new(TaRecBuf);
END; { InitPages }

BEGIN { InitTAccess }
  VSAA^.Ok := true;
  VSAA^.TAStatus := 0;
  TAErrorProc := nil;
  VSAA^.InHandler := false;
  InitPages;
  VSAA^.WarnIndex:=FALSE ;
END; { InitTAccess }

{===============================================================}
PROCEDURE FERME (sorte,typ : smallint ; reseau : boolean ) ;
BEGIN
Case typ of
 0 : closefile(VSAA^.datff[sorte]) ;
 1 : closeindex(VSAA^.idxff[sorte]) ;
 2 : BEGIN
     closefile(VSAA^.datff[sorte]) ;
     closeindex(VSAA^.idxff[sorte]) ;
     END ;
  END ;
(*???if ((Reseau=TRUE) and (Typ in [0,2])) then Libfile(sorte) ;*)
END ;


{===============================================================}
PROCEDURE OUVRE ( Sorte,Typ : smallint ; Reseau : Boolean ; Compta : TSaaProg ) ;
Var St,nomindex,nomfichier : string ;
    Rs,Ks,Dup : smallint ;
BEGIN
(*???if ((Reseau=TRUE) and (Typ in [0,2])) then Blocfile(Sorte) ;*)
St:='' ; Rs:=0 ; Ks:=0 ; Dup:=0 ;
Case Compta of
  spCPTA : BEGIN
           St:=fichiersCPTA[sorte].nom ;
           Rs:=fichiersCPTA[sorte].rec_size ;
           Ks:=fichiersCPTA[sorte].key_size ;
           Dup:=fichiersCPTA[sorte].duplicate ;
           END ;
  spNEG  : BEGIN
           St:=fichiersNEG[sorte].nom ;
           Rs:=fichiersNEG[sorte].rec_size ;
           Ks:=fichiersNEG[sorte].key_size ;
           Dup:=fichiersNEG[sorte].duplicate ;
           END ;
  spINT  : BEGIN
           St:=fichiersINT[sorte].nom ;
           Rs:=fichiersINT[sorte].rec_size ;
           Ks:=fichiersINT[sorte].key_size ;
           Dup:=fichiersINT[sorte].duplicate ;
           END ;
  END ;
if ((Sorte=7) and (Compta=spNEG)) then
   BEGIN
   Nomfichier:=SocUtil+St+'.DAT' ; NomIndex:=SocUtil+St+'.IDX' ;
   END else
   BEGIN
   Nomfichier:=SocPath+St+'.DAT' ; NomIndex:=SocPath+St+'.IDX' ;
   END ;
If (Compta=spCPTA) And (Sorte in [7,17]) Then
   BEGIN
   Nomfichier:=St+'.DAT' ; NomIndex:=St+'.IDX' ;
   END ;
Case typ of
 0 : openfile(VSAA^.datff[sorte],nomfichier,RS) ;
 1 : OpenIndex(VSAA^.idxff[sorte],NomIndex,KS,Dup) ;
 2 : BEGIN
     openfile(VSAA^.datff[sorte],nomfichier,RS) ;
     openindex(VSAA^.idxff[sorte],NomIndex,KS,Dup) ;
     END ;
 END ;
END ;


{=============================================================================}
FUNCTION LONGINT2st (Pos : Longint) : String4 ;

Var St,St1 : String4 ;

BEGIN
St:='    ' ; Move(Pos,St[1],4) ;
St1:='    ' ;
St1[1]:=St[4] ; St1[2]:=St[3] ; St1[3]:=St[2] ; St1[4]:=St[1] ;
Longint2st:=St1 ;
END ;

{============================================================================}
FUNCTION CRECLE4 ( NoIndex : smallint ; TypeE : Byte ; Journal : String3 ;
                   DateE : Word ; NoPiece : Longint ; NoOrdre : smallint ;
                   CpteGene,CpteAux : String20 ; Simul : Byte ) : String40 ;

(* EN CAS DE CHANGEMENT, PENSER A VERIFIER LA FONCTION "OPTIMISE"
   DE OUTILS !!!!! *)

Var St : String40 ;
BEGIN
St:=Chr(NoIndex) ;
Case NoIndex of
  1: BEGIN
     St:=St+Chr(TypeE) ;
     If Trim(Journal)<>'' then
        BEGIN
        St:=St+Format_String(Journal,3) ;
        If DateE<>0 then
           BEGIN
           St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
           If NoPiece<>0 then
              BEGIN
              St:=St+Longint2St(NoPiece) ;
              If NoOrdre<>0 then St:=St+Chr(Hi(NoOrdre))+Chr(Lo(NoOrdre)) ;
              END ;
           END ;
        END ;
     END ;
  2: BEGIN
     If Simul=0 then St:=St+Chr(TypeE) else St:=St+Chr(EcrReel) ;
     If trim(CpteGene)<>'' then
        BEGIN
        St:=St+Format_String(CpteGene,MaxLgCpt) ;
        If DateE<>0 then
           BEGIN
           St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
           St:=St+Chr(Simul) ;
           END ;
        END ;
     END ;
  3: BEGIN
     If Simul=0 Then St:=St+Chr(TypeE) Else St:=St+Chr(EcrReel) ;
     If trim(CpteAux)<>'' then
        BEGIN
        St:=St+Format_String(CpteAux,MaxLgCpt) ;
        If trim(CpteGene)<>'' then
           BEGIN
           St:=St+Format_String(CpteGene,MaxLgCpt) ;
           If DateE<>0 then
              BEGIN
              St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
              St:=St+Chr(Simul) ;
              END ;
           END ;
        END ;
     END ;
  END ;
CreCle4:=St ;
END ;

{============================================================================}
FUNCTION CRECLE8 ( NoIndex : smallint ; TypeE : Byte ;
                   DateE : Word ; NoPiece : Longint ; NoOrdre : smallint ;
                   NoLigne : smallint ; Section,CpteGene : String20 ; Simul : Byte ;
                   CodeJournal : String3) : String40 ;

Var St : String40 ;

BEGIN
St:=Chr(NoIndex) ;
Case NoIndex of
  1: BEGIN
     St:=St+Chr(TypeE) ;
     If trim(CodeJournal)<>'' then
        BEGIN
        St:=St+Format_String(CodeJournal,3) ;
        If DateE<>0 then
           BEGIN
           St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
           If NoPiece<>0 then
              BEGIN
              St:=St+Longint2St(NoPiece) ;
              If NoOrdre<>0 then
                 BEGIN
                 St:=St+Chr(Hi(NoOrdre))+Chr(Lo(NoOrdre)) ;
                 If NoLigne<>0 then
                    St:=St+Chr(Hi(NoLigne))+Chr(Lo(NoLigne)) ;
                 END ;
              END ;
           END ;
        END ;
     END ;
  2: BEGIN
     If Simul=0 Then St:=St+Chr(TypeE) Else St:=St+Chr(EcrReel) ;
     If trim(Section)<>'' then
        BEGIN
        St:=St+Format_String(Section,MaxLgCpt) ;
        If DateE<>0 then
           BEGIN
           St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
           If trim(CpteGene)<>'' then
              BEGIN
              St:=St+Format_String(CpteGene,MaxLgCpt) ;
              St:=St+Chr(Simul) ;
              END ;
           END ;
        END ;
     END ;
  3: BEGIN
     If Simul=0 Then St:=St+Chr(TypeE) Else St:=St+Chr(EcrReel) ;
     If trim(Section)<>'' then
        BEGIN
        St:=St+Format_String(Section,MaxLgCpt) ;
        If trim(CpteGene)<>'' then
           BEGIN
           St:=St+Format_String(CpteGene,MaxLgCpt) ;
           If DateE<>0 then
              BEGIN
              St:=St+chr(Hi(DateE))+chr(Lo(DateE)) ;
              St:=St+Chr(Simul) ;
              END ;
           END ;
        END ;
     END ;
  END ;
CreCle8:=St ;
END ;

{============================================================================}
FUNCTION CRECLE9 ( NoIndex : smallint ; TypeH : Byte ; NoPiece : Longint ;
                  NoEcrit : smallint ; NoEche,Nature : Byte ;
                  CpteGene,CpteAux : String20 ; DateH : Word ;
                  EtatH : Byte ; CodeLettreH : Longint ; DateP : Word ;
                  SimulH : Byte ) : String40 ;

Var St : String40 ;

BEGIN
St:=Chr(NoIndex) ;
Case NoIndex of
  1: BEGIN { RAJOUTER LA DATE POUR DISTINGUER 2 EXERCICES }
     St:=St+Chr(TypeH) ;
     If NoPiece<>0 then
        BEGIN
        St:=St+Longint2St(NoPiece) ;
        If NoEcrit<>0 then
           BEGIN
           St:=St+Chr(Hi(NoEcrit))+Chr(Lo(NoEcrit)) ;
           If NoEche<>0 then St:=St+Chr(NoEche) ;
           END ;
        END ;
     END ;
  2: BEGIN { Type+CpteAux+CpteGene+Etat+CodeLettre+DatePiŠce }
     If SimulH=0 then St:=St+Chr(TypeH) else St:=St+Chr(EcrReel) ;
     St:=St+Format_String(CpteAux,VSAA^.ParaSoc.FormatAux.Lg) ;
     If trim(CpteGene)<>'' then
        BEGIN
        St:=St+Format_String(CpteGene,VSAA^.ParaSoc.FormatGen.Lg) ;
        If EtatH<>0 then
           BEGIN
           St:=St+Chr(EtatH) ;
           If CodeLettreH<>-1 Then
              BEGIN
              St:=St+Longint2St(CodeLettreH) ;
              If DateH<>0 then
                 BEGIN
                 St:=St+Chr(Hi(DateH))+Chr(Lo(DateH)) ;
                 If DateP<>0 then
                    BEGIN
                    St:=St+Chr(Hi(DateP))+Chr(Lo(DateP)) ;
                    St:=St+Chr(SimulH) ;
                    END ;
                 END ;
              END ;
           END ;
        END ;
     END ;
  3: BEGIN { Type+Nature+DateEche+CpteAux+CpteGen }
     If SimulH=0 Then St:=St+Chr(TypeH) Else St:=St+Chr(EcrReel) ;
     If Nature<>0 then
        BEGIN
        St:=St+Chr(Nature) ;
        If DateP<>0 then
           BEGIN
           St:=St+Chr(Hi(DateP))+Chr(Lo(DateP)) ;
           St:=St+Format_String(CpteAux,VSAA^.ParaSoc.FormatAux.Lg) ;
           If trim(CpteGene)<>'' then
              BEGIN
              St:=St+Format_String(CpteGene,VSAA^.ParaSoc.FormatGen.Lg) ;
              St:=St+Chr(SimulH) ;
              END ;
           END ;
        END ;
     END ;
  END ;
CreCle9:=St ;
END ;


{=============================================================================}
FUNCTION CRECLE13 ( Sorte : byte ; PosFiche : Longint ; NoOrdre : smallint) : String40 ;

Var St : String40 ;

BEGIN
St:=#1 ;
St:=St+Chr(Sorte) ;
If PosFiche<>0 then
   BEGIN
   St:=St+Longint2St(PosFiche) ;
   If NoOrdre<>0 then St:=St+chr(Hi(NoOrdre))+chr(Lo(NoOrdre)) ;
   END ;
CreCle13:=St ;
END ;

{============================================================================}
FUNCTION CRECLE14 ( PosGen : Longint ; NoOrdre : SmallInt ) : String40 ;

Var St : String40 ;

BEGIN
St:=#1 ;
St:=St+Longint2St(PosGen) ;
If NoOrdre<>0 then St:=St+chr(Hi(NoOrdre))+chr(Lo(NoOrdre)) ;
CreCle14:=St ;
END ;

{=============================================================================}
FUNCTION CRECLE15 ( TypeVentil : Byte ; PosFiche : Longint ; NoOrdre : SmallInt ) : String40 ;

Var St : String40 ;

BEGIN
St:=#1 ; St:=St+Chr(TypeVentil) ;
If PosFiche<>0 then
   BEGIN
   St:=St+Longint2St(PosFiche) ;
   If NoOrdre<>0 then St:=St+chr(Hi(NoOrdre))+chr(Lo(NoOrdre)) ;
   END ;
CreCle15:=St ;
END ;

{============================================================================}
FUNCTION CRECLE16 ( NoIndex : smallint ; NoTable : Byte ; CodeTranche : String17 ;
                    Intit : String35 ; Jal : String3 ; Typ : Byte ) : String40 ;

Var St : String40 ;
    j,L : smallint ;

BEGIN
St:=Chr(NoIndex) ; St:=St+Chr(NoTable) ;
Case NoIndex Of
  1 : BEGIN
      If CodeTranche<>'' then
         BEGIN
         St:=St+CodeTranche ;
         If (NoTable in [20..39]) or (NoTable in [120..139]) or (NoTable in [140..159]) then
            BEGIN
            L:=Length(TrimRight(CodeTranche)) ; For j:=1 to MaxLgCpt-L do St:=St+'°' ;
            END ;
         END ;
      END ;
  2 : If Trim(Intit)<>'' then St:=St+Trim(Copy(Intit,1,10)) ;
  3 : If Trim(Jal)<>'' then St:=St+Trim(Jal) ;
  4 : St:=St+Chr(Typ) ;
  END ;
CreCle16:=St ;
END ;

{============================================================================}
FUNCTION CRECLEBILAN ( Sorte : Byte ; Code : String4 ) : String40 ;

Var St   : String40 ;

BEGIN
St:=#1 ; St:=St+Chr(Sorte) ;
If Trim(Code)<>'' then St:=St+Trim(Code) ;
CreCleBilan:=St ;
END ;

{============================================================================}
PROCEDURE OUVRE_BLOC ( ch : char ; Compta : TSaaProg) ;
Var SetCharRef : Set of Char ;
    FauxSocPath : String85 ;

BEGIN
SetCharRef:=['A','B','C','D','E','F','G','H','I','J'] ;
if Ch in SetCharRef then FauxSocPath:='UTIL\' else FauxSocPath:=SocPath ;
Case Compta of
  spCPTA : OpenFile(VSAA^.datfBN,FauxSocpath+'Bloc'+Ch+'.DAT',SizeOf(BNCPTA)) ;
  spNEG  : OpenFile(VSAA^.datfBN,FauxSocpath+'Bloc'+Ch+'.DAT',SizeOf(BNNEG)) ;
  END ; 
END ;

{============================================================================}
PROCEDURE FERME_BLOC ;
BEGIN
CloseFile(VSAA^.datfBN) ;
END ;


{============================================================================}
Function Int2Date ( i : word ) : TDateTime;


Var j,e,f : Word ;
    label 0 ;
Const MoisJour : array [1..13] of smallint = (0,31,59,90,120,151,181,212,243,273,304,334,366) ;

BEGIN
if i<=0 then BEGIN int2date:=EncodeDate(1900,01,01) ; goto 0 ; END ;
f:=i div 500 + 1900 ; e:=i mod 500 ;
if ((f mod 4=0) and (e>=60) and (f mod 400 <>0)) then e:=e-1 ; j:=1 ;
if e>366 then BEGIN int2date:=EncodeDate(1900,01,01) ; goto 0 ; END ;
repeat j:=j+1 ; until e<=moisjour[j] ;
j:=j-1 ; e:=i mod 500 ;
if ((f mod 4 = 0) and (e>=61) and (f mod 400 <>0)) then e:=e-moisjour[j]-1 else e:=e-moisjour[j] ;
Int2Date:=EncodeDate(f,j,e) ;
0:END ;

{$IFDEF WIN32}
{=============================================================================}
Procedure Bloc2Blob (F : TField ; B : LigBN ) ;
Var Lines : TStrings ;
    i : smallint ;
BEGIN
Lines:=TStringList.Create ;
for i:=1 to MaxligneBloc do Lines.Add(B[i]) ;
TMemoField(F).Assign(Lines) ;
Lines.Free ;
END ;



{=============================================================================}
PROCEDURE LITBLOC ( PosBloc : Longint ; Ch : Char ; Compta : TSaaProg ; T : TField) ;
Var BNC : BNCpta ;
    BNN : BNNEG ;
BEGIN
If PosBloc<>0 then
   BEGIN
   Ouvre_Bloc(ch,Compta) ;
   Case Compta of
     spCPTA : BEGIN GetRec(VSAA^.DatfBN,PosBloc,BNC) ; Bloc2Blob(T,BNC.Lignes) ; END ;
     spNEG  : BEGIN GetRec(VSAA^.DatfBN,PosBloc,BNN) ; Bloc2Blob(T,BNN.Lignes) ; END ;
     END ;
   Ferme_Bloc ;
   END ;
END ;
{$ENDIF}

{============================================================================}
PROCEDURE OUVRENUM  ;

BEGIN
AssignFile(VSAA^.FichierNum,Socpath+'Numero') ; Reset(VSAA^.FichierNum) ;
END ;

{============================================================================}
PROCEDURE FERMENUM  ;

BEGIN
Close(VSAA^.FichierNum) ;
END ;

{============================================================================}
PROCEDURE OUVRECHOIXCOD  ;
BEGIN
AssignFile(VSAA^.FichierCHCH,SocPath+'CHOIX.COD') ; Reset(VSAA^.FichierCHCH) ;
END ;

PROCEDURE FERMECHOIXCOD  ;
BEGIN
Close(VSAA^.FichierCHCH) ;
END ;

{============================================================================}
PROCEDURE OUVREDEV ;

BEGIN
Assign(VSAA^.FichierDev,Socpath+'DEVISE.COD') ; Reset(VSAA^.FichierDev) ;
END ;

{============================================================================}
PROCEDURE FERMEDEV ;

BEGIN
Close(VSAA^.FichierDev) ;
END ;

{============================================================================}
PROCEDURE OUVRE_REGLE ;

BEGIN
Assign(VSAA^.FichierRegle,Socpath+'MODEREGL.DAT') ;
{$i-} Reset(VSAA^.FichierRegle) ; {$i+} if IOresult<>0 then Rewrite(VSAA^.FichierRegle) ;
END ;

{============================================================================}
PROCEDURE FERME_REGLE ;

BEGIN
Close(VSAA^.FichierRegle) ;
END ;

{============================================================================}
PROCEDURE CHARGECALCMODE ( Num : Integer ; Var CalcMode : CalcModePaie ) ;

BEGIN
Seek(VSAA^.FichierRegle,Num-1) ; Read(VSAA^.FichierRegle,CalcMode) ;
END ;


{=============================================================================}
FUNCTION GETETAT (EtatFiche : Byte ; Etat : Byte) : Boolean ;

BEGIN
GetEtat:=(EtatFiche AND Etat)=Etat ;
END ;

{============================================================================}
PROCEDURE INITNEWSAA ( Sorte : smallint ; Var Fiche : EnregCPTA ) ;

Var oldsorte : smallint ;
    i,j : smallint ;

BEGIN
Fiche.Sorte:=Sorte ; OldSorte:=Sorte ;
if (OldSorte in [0..Nb_FichierCPTA])=FALSE then OldSorte:=0 ;
FillChar(Fiche,FichiersCPTA[OldSorte].Rec_Size,#0) ;
with Fiche Do
  Case Oldsorte of
    0 : BEGIN
        SensG:=2 ; CollectifG:=2 ; CentraliseG:=2 ; VentilableG:=2 ; PointableG:=2 ;
        LettrableG:=2 ; SoldeProgrG:=1 ; TotMensG:=2 ; SautPageG:=2 ; UniteG:=1 ;
        TvaEncaisseG:=2 ; NatureG:=9 ; PaiementG:=1 ; ModePaiementG:=1 ;  CumulFauxG:=2 ;
        SensMvtG:=1 ; EngageONG:=3 ;
        END ;
    1 : BEGIN
        NatureT:=6 ; PointableT:=2 ; SecteurT:=1 ; LettrableT:=2 ; EcheancableT:=2  ;
        RegimeTvaT:=1 ; PaiementT:=1 ; DeviseT:=1 ;RelanceT:=1 ; UniteT:=1 ;
        TvaEncaisseT:=2 ; SoldeProgrT:=1 ; TotMensT:=2 ; SautPageT:=2 ; CumulFauxT:=2 ;
        SensMvtT:=1 ; MarcheONT:=1 ;
        END ;
    2 : BEGIN
        SoldeProgrA:=1 ; TotMensA:=2 ; SautPageA:=2 ; CumulFauxA:=2 ;
        SensMvtA:=1 ;
        END ;
    3 : BEGIN
        NatureJ:=5 ; ContrepartieJ:=2 ; PieceTypeJ:=2 ; CentraliseJ:=2 ; LettrableJ:=2 ;
        AlerteCaisseJ:=2 ; CtrlTvaJ:=1 ; ChampCtrJ:=2 ; CumulFauxJ:=2 ; DeviseJ:=2 ;
        Fillchar(Fiche.OptionsPieceJ,sizeof(Fiche.OptionsPieceJ),#1) ;
        For i:=1 to 19 Do With OptionsPieceJ[i] Do
            BEGIN
            Actif:=1 ; SAisieSurListe:=2 ; RefAuto:='' ; NumSouche:='' ;
{$IFDEF STANDLITE}
            SaisieSurListe:=1 ;
{$ENDIF}
            For j:=1 To 5 Do Option[j]:=1 ;
            END ;
        For i:=1 to 18 Do With OptionsTresoJ[i] Do BEGIN AvecBor:=1 ; EcrContrep:=1 ; END ;
        END ;
    4 : BEGIN
        SensE:=1 ;UniteE:=1 ; DeviseE:=1 ; EcheanceONE:=2 ; AnalytiqueONE:=2 ;
        LettrageONE:=2 ; QuantiteONE:=2 ; EtatE:=EtatInit ; TypePieceE:=1 ;
        EncaisseE:=1 ; RegimeE:=1 ; CodeUtilE:=0 ;
        END ;
    6 : BEGIN
        TypePieceET:=1 ; TvaEncaisseET:=2 ; ContrepartieET:=2 ; TypBordET:=1 ;
        TypCtrET:=1 ; DomBanqET:=2 ; RibET:=2 ; ModReglET:=2 ; DateEcheET:=2 ;
        EtatET:=ETatInit ; PeriodiciteET:=4 ; NbRepeteET:=0 ; ReconductionET:=3 ;
        DateDepartET:=0 ; DateFinET:=0 ; NbGenereET:=0 ; LastDateET:=0 ; FinDeMoisET:=2 ;
        PosCtrPtrGET:=0 ; PosCtrPtrAET:=0 ; EcheanceONET:=1 ;
        EncaisseET:=1 ; RegimeET:=1 ; CodeUtilET:=0 ;DeviseET:=1 ; TauxDeviseET:=1 ;
        UniteET:=1 ;
        END ;
    7 : BEGIN SensRF:=1 ; CollectifRF:=2 ; END ;
    8 : BEGIN SensEA:=1 ; UniteEA:=1 ; END ;
    9 : BEGIN
        PaiementH:=1 ; SensH:=1 ; DeviseH:=1 ; EncaisseH:=1 ; RegimeH:=1 ;
        UtilisateurH:=0 ;
        END ;
{$IFDEF MAJOR}
    10 : BEGIN
         TypeEtapeER:=1 ; TypeContreER:=1 ; TypeEcritureER:=1 ; SensEcritureER:=1 ;
         TypePieceER:=3 ;
         END ;
{$ENDIF}
{   11 : BEGIN NbPeriodeC:=CurNbPeriode ; PremierePC:=CurPremierePer ; END ;}
    12 : BEGIN
         FormeJurGT:=1 ; PaysGT:=1 ; LangueGT:=1 ;
         END ;
    16 : BEGIN
         TypBordTA:=1 ; TypCtrTA:=1 ; DomBanqTA:=2 ; RibTA:=2 ; ModReglTA:=2 ; DateEcheTA:=1 ;
         END ;
    END ;
Fiche.Sorte:=Oldsorte ; Fiche.LeReseau:=0 ;
END ;

{=============================================================================}
PROCEDURE OUVREBILAN ( TT : Byte ) ;
BEGIN
If TT in [0,2] then OpenFile(VSAA^.DatFTva,SocPath+'BILAN.DAT',SizeOf(TBilan)) ;
If TT in [1,2] then OpenIndex(VSAA^.IdxFTva,SocPath+'BILAN.IDX',6,1) ;
END ;

{=============================================================================}
PROCEDURE FERMEBILAN ( TT : Byte ) ;
BEGIN
If TT in [0,2] then CloseFile(VSAA^.DatFTva) ;
If TT in [1,2] then CloseIndex(VSAA^.IdxFTva) ;
END ;

{=============================================================================}
FUNCTION CRECLEGUIDE ( Code : String ; No : SmallInt ) : String40 ;
BEGIN
CreCleGuide:=#1+Format_String(Code,8)+Chr(Hi(No))+Chr(Lo(No)) ;
END ;

{=============================================================================}
PROCEDURE OUVREGUIDE ( TT : Byte ) ;
BEGIN
if TT in [0,2] then OpenFile(VSAA^.DatFGuide,SocPath+'LIGGUIDE.DAT',SizeOf(TGuide)) ;
If TT in [1,2] then OpenIndex(VSAA^.IdxFGuide,SocPath+'LIGGUIDE.IDX',12,1) ;
END ;

{=============================================================================}
PROCEDURE FERMEGUIDE ( TT : Byte ) ;
BEGIN
If TT in [0,2] then CloseFile(VSAA^.DatFGuide)  ;
If TT in [1,2] then CloseIndex(VSAA^.IdxFGuide) ;
END ;




{===============================================================}
INITIALIZATION
InitTAccess ;

{$IFDEF WIN32}
FINALIZATION
Dispose(TaRecBuf);
Dispose(TaPgMap);
Dispose(TaPageStk);
{$A+,H+}
{$ENDIF}


END.
{$A+}

