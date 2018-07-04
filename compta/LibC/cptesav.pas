{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/07/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CPTESAV;

interface

uses SysUtils
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
  , HCtrls
  , Ent1
  , HEnt1
  , HStatus
  , Hmsgbox
  , Classes
  , Paramsoc
  , UtilSais
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  , utilPGI          // EstTablePartagee
{$IFDEF COMPTA}
  , CritEdt
  , Rapsuppr
  , SOUCHE_TOM
{$ENDIF}
  ;

type
  TLettrage = (Lettre, NonLettre, Tous);

type
  TExistMvt = record
    JOURNAL: String3;
    DATECOMPTABLE: TDateTime;
    GENERAL: String17;
    AUXIANA: String17;
    REFINTERNE: string;
    LIBELLE: string;
    DEVISE: String3;
    DEBIT: Double;
    CREDIT: Double;
    DEBITDEV: Double;
    CREDITDEV: Double;
    NUMLIGNE: Integer;
    NUMECHEVENTIL: Integer;
    ANA: String1;
    CHRONO: Integer;
  end;

function WhereCpt(Aux1, Aux2, Gen1, Gen2: string): string;
procedure SAVDATEMINDATEMAXPAQUET(Lequel: TLettrage; Aux1: string = ''; Aux2: string = ''; Gen1: string = ''; Gen2: string = '');
procedure InitExistMvt(var ExistMvt: TExistMvt);
function RecupWhereExistMvt(ExistMvt: TExistMvt): string;
function ExisteMouvement(ExistMvt: TExistMvt): boolean;
function SOLDEAZERO(FileName: string; var StErr: string): Boolean;

{$IFDEF COMPTA}
function  DetectEcritureNonValide(UnCrit: TCritEdt; LaListe: TList; NbErrD, NbErrC : THNumEdit): Integer;
function  ChercheAnaSansGene(var UnCrit: TCritEdt; LaListe: TList; OnDetruit: Boolean): Integer;
procedure RecalculSouche;
procedure RecalculSouchePourCloture;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
procedure GenereAuxSurAna(Exo: TExoDate);
{$ENDIF}

procedure RecalculTotPointe;
procedure RecalculTotPointeNew(Cpt: string);
procedure RecalculTotPointeNew1( vCompteOuJournal : string; vDossier : string = '');
procedure InitCodeAccept(WhereSup: string; Exo: TExoDate);
procedure UpdateLibelleModeleCHQ(NatModele: string);
procedure UpdateLibelleToutModele;

////////////////////////////////////////////////////////////
function CTrouveContrePartie( vJournal : string ) : string ;


implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  UTOB;

function WhereCpt(Aux1, Aux2, Gen1, Gen2: string): string;
begin
  Result := '';
  if Aux1 <> '' then
    Result := Result + 'AND E_AUXILIAIRE>="' + Aux1 + '" ';
  if Aux2 <> '' then
    Result := Result + 'AND E_AUXILIAIRE<="' + Aux2 + '" ';
  if Gen1 <> '' then
    Result := Result + 'AND E_GENERAL>="' + Gen1 + '" ';
  if Gen2 <> '' then
    Result := Result + 'AND E_GENERAL<="' + Gen2 + '" ';
end;

{=============================================================================}

procedure SAVDATEMINDATEMAXPAQUET(Lequel: TLettrage; Aux1: string = ''; Aux2:
  string = ''; Gen1: string = ''; Gen2: string = '');
var
  Q: TQuery;
  St, St1: string;
  NbMvt: Integer;
begin
  begintrans;
  InitMove(100, '');
  NbMvt := 0;
  if Lequel <> NonLettre then
  begin
    St := 'select e_auxiliaire, e_general, e_etatlettrage , e_lettrage ,min(e_datecomptable),max(e_datecomptable) from ecriture '
      +
      'where (e_etatlettrage="PL" or e_etatlettrage="TL") ';
    St1 := WhereCpt(Aux1, Aux2, Gen1, Gen2);
    if St1 <> '' then
      St := St + St1;
    St := St + ' group by e_auxiliaire, e_general, e_etatlettrage , e_lettrage';
    Q := OpenSQL(st, TRUE,-1,'',true);
    while not Q.EOF do
    begin
      MoveCur(FALSE);
      St := 'UPDATE ECRITURE set E_datepaquetmin="' +
        USDateTime(Q.Fields[4].AsDateTime) + '", ' +
        'E_DatePaquetMax="' + USDateTime(Q.Fields[5].AsDateTime) + '" ' +
        'where e_auxiliaire="' + Q.Fields[0].AsString + '" and e_general="' +
          Q.Fields[1].AsString + '" ' +
        'and e_etatlettrage="' + Q.Fields[2].AsString + '" and e_lettrage="' +
          Q.Fields[3].AsString + '"';
      ExecuteSql(St);
      Q.Next;
      if NbMvt >= 100 then
      begin
        NbMvt := 0;
        FiniMove;
        InitMove(100, '');
      end;
    end;
    Ferme(Q);
  end;
  if Lequel <> Lettre then
  begin
    MoveCur(FALSE);
    st :=
      'Update ecriture set e_datepaquetmin=e_datecomptable,e_datepaquetMax=e_datecomptable where e_etatlettrage="AL"';
    St1 := WhereCpt(Aux1, Aux2, Gen1, Gen2);
    if St1 <> '' then
      St := St + St1;
    ExecuteSQL(St);
    if Lequel = NonLettre then
      if NbMvt >= 100 then
      begin {NbMvt:=0 ;}
        FiniMove;
        InitMove(100, '');
      end;
  end;
  committrans;
  FiniMove;
end;

procedure InitExistMvt(var ExistMvt: TExistMvt);
begin
  with ExistMvt do
  begin
    JOURNAL := '';
    DATECOMPTABLE := IDate1900;
    GENERAL := '';
    AUXIANA := '';
    REFINTERNE := '';
    LIBELLE := '';
    DEVISE := '';
    DEBIT := 0;
    CREDIT := 0;
    DEBITDEV := 0;
    CREDITDEV := 0;
    NUMLIGNE := 0;
    NUMECHEVENTIL := 0;
    ANA := '-';
    CHRONO := 0;
  end;
end;

function RecupWhereExistMvt(ExistMvt: TExistMvt): string;
const
  MaxChp = 14;
var
  ChpWhere: array[1..MaxChp] of string;
  i: Integer;
  St, Pref, StWhere: string;
begin
  Result := '';
  if (ExistMvt.General = '') then
    Exit;
  FillChar(ChpWhere, SizeOf(ChpWhere), #0);
  Pref := 'E_';
  with ExistMvt do
  begin
    if JOURNAL <> '' then
      ChpWhere[1] := 'JOURNAL="' + JOURNAL;
    if (DATECOMPTABLE <> IDate1900) then
      ChpWhere[2] := 'DATECOMPTABLE="' + USDateTime(DATECOMPTABLE);
    if GENERAL <> '' then
      ChpWhere[3] := 'GENERAL="' + GENERAL;
    if AUXIANA <> '' then
      if ANA = '-' then
        ChpWhere[4] := 'AUXILIAIRE="' + AUXIANA
      else if ANA = 'X' then
        ChpWhere[4] := 'SECTION="' + AUXIANA;
    if REFINTERNE <> '' then
      ChpWhere[5] := 'REFINTERNE="' + REFINTERNE;
    if LIBELLE <> '' then
      ChpWhere[6] := 'LIBELLE="' + LIBELLE;
    if DEVISE <> '' then
      ChpWhere[7] := 'DEVISE="' + DEVISE;
    if DEBIT <> 0 then
      ChpWhere[8] := 'DEBIT=' + FloatToStr(DEBIT);
    if CREDIT <> 0 then
      ChpWhere[9] := 'CREDIT=' + FloatToStr(CREDIT);
    if DEBITDEV <> 0 then
      ChpWhere[10] := 'DEBITDEV=' + FloatToStr(DEBITDEV);
    if CREDITDEV <> 0 then
      ChpWhere[10] := 'CREDITDEV=' + FloatToStr(CREDITDEV);
    if NUMLIGNE <> 0 then
      ChpWhere[11] := 'NUMLIGNE=' + IntToStr(NUMLIGNE);
    if CHRONO <> 0 then
      ChpWhere[14] := 'CHRONO<>' + IntToStr(CHRONO);
    if ANA <> '' then
    begin
      if ANA = '-' then
      begin
        if NUMECHEVENTIL <> 0 then
          ChpWhere[12] := 'NUMECHE=' + IntToStr(NUMECHEVENTIL);
      end
      else if ANA = 'X' then
      begin
        ChpWhere[13] := 'ANA="' + ANA;
        Pref := 'Y_';
        if NUMECHEVENTIL <> 0 then
          ChpWhere[12] := 'NUMVENTIL=' + IntToStr(NUMECHEVENTIL);
      end;
    end;
  end;

  for i := 1 to MaxChp do
  begin
    if (ChpWhere[i] <> '') then
    begin
      if (StWhere <> '') then
        StWhere := StWhere + ' AND ';
      St := Pref + ChpWhere[i];
      case (Pos('"', St) = 0) of
        True: if (Pos(',', St) <> 0) then
            St[Pos(',', St)] := '.';
        False: St := St + '"';
      end;
      StWhere := StWhere + St;
    end;
  end;
  Result := StWhere;
end;

function ExisteMouvement(ExistMvt: TExistMvt): boolean;
var
  Q: TQuery;
  SQL, Pref, Tabl: string;
begin
  Result := False;
  if ExistMvt.General = '' then
    Exit;
  if ExistMvt.ANA = 'X' then
  begin
    Pref := 'Y_';
    Tabl := 'ANALYTIQ';
  end
  else
  begin
    Pref := 'E_';
    Tabl := 'ECRITURE';
  end;
  SQL := 'SELECT G_GENERAL FROM GENERAUX WHERE EXISTS(SELECT ' + Pref +
    'GENERAL FROM ' + Tabl + ' WHERE ';
  SQL := SQL + RecupWhereExistMvt(ExistMvt) + ') AND G_GENERAL="' +
    ExistMvt.General + '"';
  Q := OpenSQL(SQL, True,-1,'',true);
  Result := not Q.Eof;
  Ferme(Q);
end;

function FORMAT1_STRING(st: string; l: Integer): string;
var
  St1: string;
begin
  St1 := St;
  if ((l > 0) and (l < Length(St1))) then
    St1 := Copy(St1, 1, l);
  while Length(St1) < l do
    St1 := St1 + ' ';
  Result := st1;
end;

function VerifMontant(St: string): Boolean;
var
  i: Integer;
begin
  Result := TRUE;
  for i := 1 to length(St) do
  begin
    if (St[i] in ['0'..'9', '.', ',', '-']) = FALSE then
    begin
      Result := FALSE;
      Exit;
    end;
  end;
end;

function VerifEntier(St: string): Boolean;
var
  i: Integer;
begin
  Result := TRUE;
  for i := 1 to length(St) do
  begin
    if (St[i] in ['0'..'9']) = FALSE then
    begin
      Result := FALSE;
      Exit;
    end;
  end;
end;

function SOLDEAZERO(FileName: string; var StErr: string): Boolean;

var
  Fichier, NewFichier: TextFile;
  NewFileName: string;
  St: string;
  StTest, Sens: string;
  NoPiece, PieceFictive: Integer;
  TotPieceD, TotPieceC: Double;
  Montant: Double;
  Longueur: integer;
  OkMontant, OkNumP: Boolean;
  NumLigne: Integer;
begin
  Result := TRUE;
  StErr := '';
  Longueur := 132;
  AssignFile(Fichier, FileName);
  Reset(Fichier);
  NewFileName := FileTemp('.PNM');
  AssignFile(NewFichier, NewFileName);
  Rewrite(NewFichier);
  ReadLn(Fichier, St);
  St := Format_String(St, Longueur);
  WriteLn(NewFichier, St);
  PieceFictive := 0;
  TotPieceD := 0;
  TotPieceC := 0;
  NumLigne := 1;
  while not EOF(Fichier) do
  begin
    Inc(NumLigne);
    ReadLn(Fichier, St);
    if Length(St) = 467 then
      Longueur := 467; // Ajout des infos comptes en EOR.
    St := Format1_String(St, Longueur);
    if Trim(St) <> '' then
    begin
      StTest := Trim(Copy(St, 106, 7));
      OkNumP := VerifEntier(StTest);
      if (StTest <> '') and OkNumP then
        NoPiece := StrToInt(StTest)
      else
        NoPiece := 0;
      StTest := Trim(Copy(St, 85, 20));
      OkMontant := VerifMontant(StTest);
      if (StTest <> '') and OkMontant then
        Montant := StrToFloat(StPoint(StTest))
      else
        Montant := 0;
      if (not OkMontant) or (not OkNumP) then
      begin
        StErr := StErr + IntToStr(NumLigne) + ';';
        Result := FALSE;
      end;
      Sens := Copy(St, 84, 1);
      if (NoPiece = 0) and OkNumP then
      begin
        if (St[25] <> 'A') and (St[25] <> 'E') and (St[25] <> 'D') then
          { pour shunter les écritures analytiques }
        begin
          if (Arrondi(TotPieceD, V_PGI.OkDecV) = Arrondi(TotPieceC,
            V_PGI.OkDecV)) then
          begin
            Inc(PieceFictive);
            TotPieceD := 0;
            TotPieceC := 0;
          end;
          if Sens = 'D' then
            TotPieceD := Arrondi(TotPieceD + Montant, V_PGI.OkDecV)
          else
            TotPieceC := Arrondi(TotPieceC + Montant, V_PGI.OkDecV);
        end;
        St := Insere(St, FormatFloat('0000000', PieceFictive), 106, 7);
      end;
      WriteLn(NewFichier, St);
    end;
  end;
  CloseFile(Fichier);
  CloseFile(NewFichier);
  AssignFile(Fichier, FileName);
  Erase(Fichier);
  renamefile(NewFileName, FileName);
end;

{$IFDEF COMPTA}
// FQ 10394
function DetectEcritureNonValide(UnCrit: TCritEdt; LaListe: TList; NbErrD, NbErrC : THNumEdit): Integer;
var
  QLoc: TQuery;
  StSelect, Sql, SqlWhere: string;
  MemoNumPiece: Integer;
  bOkCorrige : Boolean;

  Function GoListe1(Q : TQuery) : DelInfo ;
  var
    X : DelInfo ;
  BEGIN
  X:=DelInfo.Create ;
  X.LeCod := QLoc.FindField('E_JOURNAL').AsString;
  X.LeLib := QLoc.FindField('E_REFINTERNE').AsString;
  X.LeMess:=QLoc.FindField('E_NUMEROPIECE').AsString + '/' + QLoc.FindField('E_NUMLIGNE').AsString;
  X.LeMess2:=QLoc.FindField('E_DATECOMPTABLE').AsString;
  X.LeMess3:='Le champ E_QUALIFPIECE est incorrect';
  X.LeMess4:=QLoc.FindField('E_EXERCICE').AsString;
  Result:=X ;
  END ;

begin
  Result := 0;
  StSelect := 'Select * From ECRITURE Where E_JOURNAL in (Select J_JOURNAL From JOURNAL Where '
   +'J_NATUREJAL="BQE" Or J_NATUREJAL="CAI") And ' +
    '(E_GENERAL in (Select G_GENERAL From GENERAUX Where G_LETTRABLE="X") Or ' +
    'E_AUXILIAIRE in (Select T_AUXILIAIRE From TIERS Where T_LETTRABLE="X"))';

  SqlWhere := 'And E_QUALIFPIECE<>"N" And E_QUALIFPIECE<>"S" And ' +
    'E_QUALIFPIECE<>"U" And E_QUALIFPIECE<>"R" And E_QUALIFPIECE<>"P" And ' +
    'E_QUALIFPIECE<>"C" And E_EXERCICE="' + UnCrit.Exo.Code + '" And ' +
    'E_DATECOMPTABLE>="' + UsDateTime(UnCrit.Date1) + '" And E_DATECOMPTABLE<="'
      + UsDateTime(UnCrit.Date2) + '" And ' +
    'E_NUMEROPIECE>=' + IntToStr(UnCrit.GL.NumPiece1) + ' And E_NUMEROPIECE<=' +
      IntTostr(UnCrit.GL.NumPiece2) + '';
  if UnCrit.Etab <> '' then
    SqlWhere := SqlWhere + ' And E_ETABLISSEMENT="' + UnCrit.Etab + '"';
  if UnCrit.Cpt1 <> '' then
    SqlWhere := SqlWhere + ' And E_JOURNAL>="' + UnCrit.Cpt1 + '"';
  if UnCrit.Cpt2 <> '' then
    SqlWhere := SqlWhere + ' And E_JOURNAL<="' + UnCrit.Cpt2 + '"';
  Sql := StSelect + SqlWhere;
  QLoc := OpenSql(Sql, True,-1,'',true);
  MemoNumPiece := 0;
  while not QLoc.Eof do
  begin
    if MemoNumPiece = QLoc.FindField('E_NUMEROPIECE').AsInteger then
    begin
      QLoc.Next;
      Continue;
    end;
    try
      LaListe.Add(GoListe1(QLoc));
      NbErrD.Value := NbErrD.Value+1;
      inc(Result);
      bOkCorrige := True;
      BeginTrans;
      ExecuteSql('Update ECRITURE Set E_TRESOLETTRE="-",E_QUALIFPIECE="N" Where '
        +
        'E_EXERCICE="' + UnCrit.Exo.Code + '" And ' +
        'E_NUMEROPIECE=' + IntToStr(QLoc.FindField('E_NUMEROPIECE').AsInteger) +
          ' And ' +
        'E_ETABLISSEMENT="' + QLoc.FindField('E_ETABLISSEMENT').AsString +
          '" And ' +
        'E_JOURNAL="' + QLoc.FindField('E_JOURNAL').AsString + '"');
      CommitTrans;
    except
      bOkCorrige := False;
      HShowMessage('0;Réparation comptable;La mise à jour ne peut être effectuée. La pièce va être détruite;W;O;O;O;', '', '');
      Rollback;
      BeginTrans;
      MemoNumPiece := QLoc.FindField('E_NUMEROPIECE').AsInteger;
      ExecuteSql('Delete From ECRITURE Where E_JOURNAL="' +
        QLoc.FindField('E_JOURNAL').AsString + '" And ' +
        'E_EXERCICE="' + UnCrit.Exo.Code + '" And ' +
        'E_NUMEROPIECE=' + IntToStr(QLoc.FindField('E_NUMEROPIECE').AsInteger) +
          ' And ' +
        'E_ETABLISSEMENT="' + QLoc.FindField('E_ETABLISSEMENT').AsString + '"');
      CommitTrans;
    end;
    if bOkCorrige then NbErrC.Value := NbErrC.Value+1;
    QLoc.Next;
  end;
  Ferme(QLoc);

  SqlWhere := 'And E_TRESOLETTRE<>"-" And ' +
    'E_QUALIFPIECE="' + UnCrit.QualifPiece + '" And E_EXERCICE="' +
      UnCrit.Exo.Code + '" And ' +
    'E_DATECOMPTABLE>="' + UsDateTime(UnCrit.Date1) + '" And E_DATECOMPTABLE<="'
      + UsDateTime(UnCrit.Date2) + '" And ' +
    'E_NUMEROPIECE>=' + IntToStr(UnCrit.GL.NumPiece1) + ' And E_NUMEROPIECE<=' +
      IntTostr(UnCrit.GL.NumPiece2) + '';
  if UnCrit.Etab <> '' then
    SqlWhere := SqlWhere + ' And E_ETABLISSEMENT="' + UnCrit.Etab + '"';
  if UnCrit.Cpt1 <> '' then
    SqlWhere := SqlWhere + ' And E_JOURNAL>="' + UnCrit.Cpt1 + '"';
  if UnCrit.Cpt2 <> '' then
    SqlWhere := SqlWhere + ' And E_JOURNAL<="' + UnCrit.Cpt2 + '"';
  Sql := StSelect + SqlWhere;
  QLoc := OpenSql(Sql, True,-1,'',true);
  while not QLoc.Eof do
  begin
    LaListe.Add(GoListe1(QLoc));
    inc(Result);
    NbErrD.Value := NbErrD.Value+1;
    NbErrC.Value := NbErrC.Value+1;
    BeginTrans;
    ExecuteSql('Update ECRITURE Set E_TRESOLETTRE="-" Where ' +
      'E_EXERCICE="' + UnCrit.Exo.Code + '" And ' +
      'E_NUMEROPIECE=' + IntToStr(QLoc.FindField('E_NUMEROPIECE').AsInteger) +
        ' And ' +
      'E_ETABLISSEMENT="' + QLoc.FindField('E_ETABLISSEMENT').AsString + '" And '
        +
      'E_QUALIFPIECE="' + UnCrit.QualifPiece + '" And ' +
      'E_JOURNAL="' + QLoc.FindField('E_JOURNAL').AsString + '"');
    CommitTrans;
    QLoc.Next;
  end;
  Ferme(QLoc);
end;

function ChercheAnaSansGene(var UnCrit: TCritEdt; LaListe: TList; OnDetruit: Boolean): Integer;
var
  Sql, StSelect, StWhere: string;
  QLoc: TQuery;
  X: DelInfo;
begin
  Result := 0;
  try
    BeginTrans;
    // If OnDetruit Then LaListe.Clear ;
    StSelect :=
      'SELECT Y_JOURNAL,Y_EXERCICE,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_NUMVENTIL,Y_QUALIFPIECE,Y_GENERAL,Y_SECTION,Y_REFINTERNE,Y_AXE,Y_ETABLISSEMENT FROM ANALYTIQ Q ';
    StWhere := 'Where Y_EXERCICE="' + UnCrit.Exo.Code + '" And ' +
      'Y_DATECOMPTABLE>="' + UsDateTime(UnCrit.Date1) +
        '" And Y_DATECOMPTABLE<="' + UsDateTime(UnCrit.Date2) + '" And ' +
      'Y_NUMEROPIECE>=' + IntToStr(UnCrit.GL.NumPiece1) + ' And Y_NUMEROPIECE<='
        + IntTostr(UnCrit.GL.NumPiece2) + ' ';
    if UnCrit.Etab <> '' then
      StWhere := StWhere + ' And Y_ETABLISSEMENT="' + UnCrit.Etab + '"';
    if UnCrit.Cpt1 <> '' then
      StWhere := StWhere + ' And Y_JOURNAL>="' + UnCrit.Cpt1 + '"';
    if UnCrit.Cpt2 <> '' then
      StWhere := StWhere + ' And Y_JOURNAL<="' + UnCrit.Cpt2 + '"';
    StWhere := StWhere +
      ' And Y_TYPEANALYTIQUE="-" And Y_QUALIFPIECE<>"C" And Y_ECRANOUVEAU="N"';
    StWhere := StWhere +
      ' AND NOT EXISTS(SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_QUALIFPIECE FROM ECRITURE ';
    StWhere := StWhere +
      'WHERE E_JOURNAL=Q.Y_JOURNAL and E_EXERCICE=Q.Y_EXERCICE AND E_DATECOMPTABLE=Q.Y_DATECOMPTABLE and E_NUMEROPIECE=Q.Y_NUMEROPIECE AND ';
    StWhere := StWhere +
      'E_NUMLIGNE=Q.Y_NUMLIGNE AND E_QUALIFPIECE=Q.Y_QUALIFPIECE AND E_GENERAL=Q.Y_GENERAL)';
    Sql := StSelect + StWhere;
    QLoc := OpenSql(Sql, True,-1,'',true);
    while not QLoc.Eof do
    begin
      Inc(Result);
      if LaListe <> nil then
      begin
        X := DelInfo.Create;
        X.Gen := QLoc.FindField('Y_GENERAL').AsString;
        X.Sect := QLoc.FindField('Y_SECTION').AsString;
        X.LeCod := QLoc.FindField('Y_JOURNAL').AsString;
        X.LeLib := QLoc.FindField('Y_REFINTERNE').AsString;
        X.LeMess := IntToStr(QLoc.FindField('Y_NUMEROPIECE').AsInteger) + '/' +
          IntToStr(QLoc.FindField('Y_NUMLIGNE').AsInteger);
        X.LeMess2 := DateToStr(QLoc.FindField('Y_DATECOMPTABLE').AsDateTime);
        X.LeMess3 :=
          'Analytique : Ce mouvement n''a pas d''écriture comptable rattachée (Section : ' +
          X.Sect + ')';
        X.LeMess4 := QLoc.FindField('Y_EXERCICE').AsString + ';' +
          QLoc.FindField('Y_AXE').AsString;
        LaListe.Add(X);
      end;
      if OnDetruit then
      begin
        (*
        X.LeLib:=QLoc.FindField('Y_NUMEROPIECE').AsString ;
        X.LeMess:=QLoc.FindField('Y_DATECOMPTABLE').AsString ;
        X.LeMess2:=('Pièce analytique sans écriture comptable reliée') ;
        X.LeMess2:=('Pièce analytique détruite') ;
        *)
        ExecuteSql('Delete From ANALYTIQ Where Y_JOURNAL="' +
          QLoc.FindField('Y_JOURNAL').AsString + '" And ' +
          'Y_NUMEROPIECE=' + IntToStr(QLoc.FindField('Y_NUMEROPIECE').AsInteger)
            + ' And ' +
          'Y_DATECOMPTABLE="' +
            USDATETIME(QLoc.FindField('Y_DATECOMPTABLE').AsDateTime) + '" And ' +
          'Y_EXERCICE="' + QLoc.FindField('Y_EXERCICE').AsString + '" And ' +
          'Y_ETABLISSEMENT="' + QLoc.FindField('Y_ETABLISSEMENT').AsString +
            '" And ' +
          'Y_NUMLIGNE=' + IntToStr(QLoc.FindField('Y_NUMLIGNE').AsInteger) +
            ' AND ' +
          'Y_QUALIFPIECE="' + QLoc.FindField('Y_QUALIFPIECE').AsString + '" ');
      end;
      //    If LaListe<>Nil Then LaListe.Add(X) ;
      QLoc.Next;
    end;
    (*
    if (LaListe.Count>0) And OnDetruit then
       HShowMessage('0;Réparation comptable;Les pièces analytiques sans écriture comptable reliée ont été détruites;W;O;O;O;','','') ;
       else if LaListe.Count>0 then
            HShowMessage('0;Vérification comptable;Il y a des pièces analytiques sans écriture comptable reliée;W;O;O;O;','','') ;
    *)
    Ferme(QLoc);
    CommitTrans;
  except
    Rollback;
  end;
end;

{$ENDIF}

{$IFNDEF EAGLCLIENT}
procedure GenereAuxSurAna(Exo: TExoDate);
var
  Q, Q1: TQuery;
  St, StSelect, StWhere, StSQL: string;
  PremFois: Boolean;
begin
  St := '';
  Q :=
    OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTIF="X" AND G_VENTILABLE="X"', TRUE,-1,'',true);
  PremFois := TRUE;
  while not Q.Eof do
  begin
    if PremFois then
      St := 'E_GENERAL="' + Q.Fields[0].AsString + '" '
    else
      St := St + ' Or E_GENERAL="' + Q.Fields[0].AsString + '" ';
    PremFois := FALSE;
    Q.Next;
  end;
  Ferme(Q);
  if St = '' then
    Exit;
  try
    BeginTrans;
    St := ' (' + St + ') ';
    StSelect :=
      'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_QUALIFPIECE,E_AUXILIAIRE FROM ECRITURE ';
    StWhere := 'Where E_EXERCICE="' + Exo.Code + '" And E_DATECOMPTABLE>="' +
      UsDateTime(Exo.Deb) + '" And E_DATECOMPTABLE<="' + UsDateTime(Exo.Fin) +
      '" ';
    StWhere := StWhere + ' And E_QUALIFPIECE="N" ';
    StWhere := StWhere + ' AND ' + St;
    StSql := StSelect + StWhere;
    Q := OpenSql(StSql, True,-1,'',true);
    StSelect := 'SELECT Y_AUXILIAIRE FROM ANALYTIQ ';
    StWhere :=
      'Where Y_EXERCICE=:EXO And Y_DATECOMPTABLE=:DATE AND Y_JOURNAL=:JAL AND Y_NUMEROPIECE=:NUMP AND Y_NUMLIGNE=:NUML AND Y_QUALIFPIECE=:QUALP';
    StSql := StSelect + StWhere;
    Q1 := PrepareSQL(StSql, FALSE);
    InitMove(RecordsCount(Q), '');
    while not Q.Eof do
    begin
      MoveCur(FALSE);
      Q1.Close;
      Q1.Params[0].AsString := Q.FindField('E_EXERCICE').AsString;
      Q1.Params[1].AsDateTime := Q.FindField('E_DATECOMPTABLE').AsDateTime;
      Q1.Params[2].AsString := Q.FindField('E_JOURNAL').AsString;
      Q1.Params[3].AsInteger := Q.FindField('E_NUMEROPIECE').AsInteger;
      Q1.Params[4].AsInteger := Q.FindField('E_NUMLIGNE').AsInteger;
      Q1.Params[5].AsString := Q.FindField('E_QUALIFPIECE').AsString;
      Q1.Open;
      while not Q1.EOF do
      begin
        Q1.Edit;
        Q1.Fields[0].AsString := Q.FindField('E_AUXILIAIRE').AsString;
        Q1.Post;
        Q1.Next;
      end;
      Q.Next;
    end;
    Ferme(Q1);
    Ferme(Q);
    FiniMove;
    CommitTrans;
  except
    Rollback;
  end;
end;
{$ENDIF}

{$IFDEF COMPTA}
procedure RecalculSouche;
var
  Sql, Jal, Souche, StWhereJal: string;
  Q, Q1: TQuery;
  OkN1: Boolean;
  Exo: TExoDate;
begin
  try
    BeginTrans;
    SQL :=
      'Select * From Souche Where SH_TYPE="CPT" And SH_ANALYTIQUE="-" AND SH_SIMULATION="-"';
    Q := OpenSql(Sql, True,-1,'',true);
    while not Q.Eof do
    begin
      Souche := Q.FindField('SH_SOUCHE').AsString;
      Fillchar(Exo, SizeOf(Exo), #0);
      OkN1 := VH^.MultiSouche and (Q.FindField('SH_SOUCHEEXO').AsString = 'X');
      SQL := 'Select J_JOURNAL from JOURNAL Where J_COMPTEURNORMAL="' + Souche +
        '"';
      Q1 := OpenSql(Sql, True,-1,'',true);
//      FetchSQLODBC(Q1);
      StWhereJal := '';
      while not Q1.Eof do
      begin
        Jal := Q1.FindField('J_JOURNAL').AsString;
        if StWhereJal = '' then
          StWhereJal := 'E_JOURNAL="' + Jal + '" '
        else
          StWhereJal := StWhereJal + ' OR E_JOURNAL="' + Jal + '" ';
        Q1.Next;
      end;
      Ferme(Q1);
      if StWhereJal <> '' then
      begin
        StWhereJal := '(' + StWhereJal + ') ';
        if VH^.MultiSouche and OkN1 then
        begin
          RecalculSOucheSurUnExo(StWhereJal, Souche, VH^.Precedent, TRUE);
          RecalculSOucheSurUnExo(StWhereJal, Souche, VH^.EnCours, TRUE);
          RecalculSOucheSurUnExo(StWhereJal, Souche, VH^.Suivant, TRUE);
        end
        else
          RecalculSOucheSurUnExo(StWhereJal, Souche, Exo, TRUE);
        (*
              Num:=0 ;
              SQL:='Select Max(E_NUMEROPIECE) From ECRITURE Where '+StWhereJal+' AND E_QUALIFPIECE="N"' ;
              Q1:=OpenSql(Sql,True) ; If Not Q1.Eof Then Num:=Q1.Fields[0].AsInteger ;
              Ferme(Q1) ;
              Inc(Num) ;
              SQL:='Update SOUCHE Set SH_NUMDEPART='+IntToStr(Num)+'WHERE SH_SOUCHE="'+Souche+'"' ;
              ExecuteSQL(SQL) ;
        *)
      end;
      Q.Next;
    end;
    Ferme(Q);
    CommitTrans;
  except
    Rollback;
  end;
end;

procedure RecalculSouchePourCloture;
var
  Sql, Jal, Souche, StWhereJal: string;
  Q, Q1: TQuery;
  Exo: TExoDate;
begin
  if not VH^.MultiSouche then
    Exit;
  try
    BeginTrans;
    SQL :=
      'Select * From Souche Where SH_TYPE="CPT" And SH_ANALYTIQUE="-" AND SH_SIMULATION="-" AND SH_SOUCHEEXO="X"';
    Q := OpenSql(Sql, True,-1,'',true);
    while not Q.Eof do
    begin
      Souche := Q.FindField('SH_SOUCHE').AsString;
      Fillchar(Exo, SizeOf(Exo), #0);
      SQL := 'Select J_JOURNAL from JOURNAL Where J_COMPTEURNORMAL="' + Souche +
        '"';
      Q1 := OpenSql(Sql, True,-1,'',true);
//      FetchSQLODBC(Q1);
      StWhereJal := '';
      while not Q1.Eof do
      begin
        Jal := Q1.FindField('J_JOURNAL').AsString;
        if StWhereJal = '' then
          StWhereJal := 'E_JOURNAL="' + Jal + '" '
        else
          StWhereJal := StWhereJal + ' OR E_JOURNAL="' + Jal + '" ';
        Q1.Next;
      end;
      Ferme(Q1);
      if StWhereJal <> '' then
      begin
        StWhereJal := '(' + StWhereJal + ') ';
        RecalculSoucheSurUnExo(StWhereJal, Souche, VH^.Precedent, TRUE);
        RecalculSoucheSurUnExo(StWhereJal, Souche, VH^.EnCours, TRUE);
        RecalculSoucheSurUnExo(StWhereJal, Souche, VH^.Suivant, TRUE);
      end;
      Q.Next;
    end;
    Ferme(Q);
    CommitTrans;
  except
    Rollback;
  end;
end;
{$ENDIF}

procedure RecalculTotPointe;
var
  St, StSup: string;
  Q, Q1: TQuery;
begin
  St := 'UPDATE GENERAUX SET G_TOTDEBPTP=0, G_TOTCREPTP=0, G_TOTDEBPTD=0, G_TOTCREPTD=0';
  ExecuteSQL(St);
  CReinitCumulsPointeMS ;

  if VH^.Suivant.Code <> '' then
    StSup := '(E_EXERCICE="' + VH^.EnCours.Code + '" OR E_EXERCICE="' + VH^.Suivant.Code + '") '
  else
    StSup := 'E_EXERCICE="' + VH^.EnCours.Code + '" ';

  St := 'SELECT E_GENERAL, SUM(E_DEBIT), SUM(E_CREDIT), SUM(E_DEBITDEV), SUM(E_CREDITDEV) FROM ECRITURE ' +
        'LEFT JOIN GENERAUX on G_GENERAL=E_GENERAL WHERE G_NATUREGENE="BQE" AND G_POINTABLE="X" AND ' +
        'E_REFPOINTAGE<>"" AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="OAN") AND ' +
        StSup + ' GROUP BY E_GENERAL ';

  Q := OpenSQL(St, TRUE,-1,'',true);
  while not Q.Eof do
  begin
    St := 'Select G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD ' +
          'FROM GENERAUX WHERE G_GENERAL="' + Q.Fields[0].AsString + '" ';
    Q1 := OpenSQL(St, FALSE);
    if not Q1.EOF then
    begin
      if not EstTablePartagee( 'GENERAUX' ) then
        begin
        Q1.Edit;
        Q1.FindField('G_TOTDEBPTP').AsFloat := Q.Fields[1].AsFloat;
        Q1.FindField('G_TOTCREPTP').AsFloat := Q.Fields[2].AsFloat;
        Q1.FindField('G_TOTDEBPTD').AsFloat := Q.Fields[3].AsFloat;
        Q1.FindField('G_TOTCREPTD').AsFloat := Q.Fields[4].AsFloat;
        Q1.Post;
        end
      else
        // MAJ table cumuls si besoin
        CUpdateCumulsPointeMS( Q.Fields[0].AsString, Q.Fields[1].AsFloat, Q.Fields[2].AsFloat, Q.Fields[3].AsFloat, Q.Fields[4].AsFloat ) ;
    end;
    Ferme(Q1);
    Q.Next;
  end;
  Ferme(Q);
end;

procedure CalcTotPointeExo(AvantEnCours: Boolean);
{Var St,StSup,StSup1 : String ;
    Q,Q1,Q2 : TQuery ;
    OldGen,CodeDev : String ;
    StCpt : String ;}
begin
  (*
  If AvantEnCours Then
    BEGIN
    StSup:='AND E_DATECOMPTABLE<"'+USDateTime(VH^.EnCours.code)+'" ' ;
    if VH^.ExoV8.Code<>'' then StSup1:='E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" '

    StSup1:=' AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="OAN") ' ;
    END Else
    BEGIN
    If VH^.Suivant.Code<>'' Then StSup:='AND (E_EXERCICE="'+VH^.EnCours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") '
                            Else StSup:='AND E_EXERCICE="'+VH^.EnCours.Code+'" ' ;
    StSup1:=' AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="OAN") ' ;
    END ;
  If StCpt<>'' Then StCpt:='AND'+StCpt ;
  St:='SELECT E_GENERAL, E_DEVISE, SUM(E_DEBIT), SUM(E_CREDIT), SUM(E_DEBITDEV), SUM(E_CREDITDEV), '+
      'FROM ECRITURE '+
      'LEFT JOIN GENERAUX on G_GENERAL=E_GENERAL WHERE G_NATUREGENE="BQE" AND G_POINTABLE="X" AND '+
      'E_REFPOINTAGE<>"" AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="OAN") '+
      StSup+StCpt+' GROUP BY E_GENERAL,E_DEVISE ' ;
  St:=St+' ORDER BY E_GENERAL,E_DEVISE ' ;
  Q:=OpenSQL(St,TRUE) ; OldGen:='' ;
  FetchSQLODBC(Q) ;
  InitMove(RecordsCount(Q),'') ;
  While Not Q.Eof Do
    BEGIN
    MoveCur(FALSE) ;
    St:='Select G_TOTDEBPTP,G_TOTCREPTP,G_TOTDEBPTD,G_TOTCREPTD '+
        'FROM GENERAUX WHERE G_GENERAL="'+Q.Fields[0].AsString+'" ' ;
    Q1:=OpenSQL(St,FALSE) ;
    If Not Q1.EOF Then
      BEGIN
      If OldGen<>Q.Fields[0].AsString Then
        BEGIN
        Q2:=OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL="'+Q.Fields[0].AsString+'" ',TRUE) ;
        If Not Q2.Eof Then CodeDev:=Q2.Fields[0].AsString ;
        Ferme(Q2) ;
        OldGen:=Q.Fields[0].AsString ;
        END ;
      Q1.Edit ;
      Q1.FindField('G_TOTDEBPTP').AsFloat:=Q1.FindField('G_TOTDEBPTP').AsFloat+Q.Fields[2].AsFloat ;
      Q1.FindField('G_TOTCREPTP').AsFloat:=Q1.FindField('G_TOTCREPTP').AsFloat+Q.Fields[3].AsFloat ;
      If CodeDev=Q.Fields[1].AsString Then
        BEGIN
        Q1.FindField('G_TOTDEBPTD').AsFloat:=Q1.FindField('G_TOTDEBPTD').AsFloat+Q.Fields[4].AsFloat ;
        Q1.FindField('G_TOTCREPTD').AsFloat:=Q1.FindField('G_TOTCREPTD').AsFloat+Q.Fields[5].AsFloat ;
        END ;
      Q1.Post ;
      END ;
    Ferme(Q1) ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  FiniMove ;
  *)
end;

procedure RecalculTotPointeNew(Cpt: string);
var
  St, StSup: string;
  Q, Q1, Q2: TQuery;
  OldGen, CodeDev: string;
  StCpt: string;
begin
  StCpt := '';
  if Cpt <> '' then
    StCpt := ' G_GENERAL="' + Cpt + '" ';

  // initialisation
  St := 'UPDATE GENERAUX SET G_TOTDEBPTP = 0, G_TOTCREPTP = 0, G_TOTDEBPTD = 0, G_TOTCREPTD = 0';
  if StCpt <> '' then
    St := St + 'WHERE' + StCpt;
  ExecuteSQL(St);
  CReinitCumulsPointeMS( Cpt ) ;

  // Recalcul
  if VH^.Suivant.Code <> '' then
    StSup := 'AND (E_EXERCICE="' + VH^.EnCours.Code + '" OR E_EXERCICE="' + VH^.Suivant.Code + '") '
  else
    StSup := 'AND E_EXERCICE="' + VH^.EnCours.Code + '" ';

  if StCpt <> '' then
    StCpt := 'AND' + StCpt;
    St := 'SELECT E_GENERAL, E_DEVISE, SUM(E_DEBIT), SUM(E_CREDIT), SUM(E_DEBITDEV), SUM(E_CREDITDEV), ' +
    ' FROM ECRITURE ' +
    'LEFT JOIN GENERAUX on G_GENERAL=E_GENERAL WHERE G_NATUREGENE="BQE" AND G_POINTABLE="X" AND ' +
    'E_REFPOINTAGE<>"" AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="OAN") ' +
    StSup + StCpt + ' GROUP BY E_GENERAL,E_DEVISE ';

  St := St + ' ORDER BY E_GENERAL,E_DEVISE ';
  Q := OpenSQL(St, TRUE);
  OldGen := '';
//  FetchSQLODBC(Q);
  InitMove(RecordsCount(Q), '');
  while not Q.Eof do
  begin
    MoveCur(FALSE);
    St := 'Select G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD FROM GENERAUX ' +
          'WHERE G_GENERAL = "' + Q.Fields[0].AsString + '" ';
    Q1 := OpenSQL(St, FALSE);
    if not Q1.EOF then
    begin
      if OldGen <> Q.Fields[0].AsString then
      begin
        Q2 := OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL="' + Q.Fields[0].AsString
        + '" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ', True); // 24/10/2006 YMO Multisociétés
        if not Q2.Eof then
          CodeDev := Q2.Fields[0].AsString;
        Ferme(Q2);
        OldGen := Q.Fields[0].AsString;
      end;
      Q1.Edit;
      Q1.FindField('G_TOTDEBPTP').AsFloat := Q1.FindField('G_TOTDEBPTP').AsFloat + Q.Fields[2].AsFloat;
      Q1.FindField('G_TOTCREPTP').AsFloat := Q1.FindField('G_TOTCREPTP').AsFloat + Q.Fields[3].AsFloat;

      if CodeDev = Q.Fields[1].AsString then
      begin
        Q1.FindField('G_TOTDEBPTD').AsFloat := Q1.FindField('G_TOTDEBPTD').AsFloat + Q.Fields[4].AsFloat;
        Q1.FindField('G_TOTCREPTD').AsFloat := Q1.FindField('G_TOTCREPTD').AsFloat + Q.Fields[5].AsFloat;
      end;
      Q1.Post;

      // MAJ table cumuls si besoin
      if CodeDev = Q.Fields[1].AsString
        then CUpdateCumulsPointeMS( Q.Fields[0].AsString, Q.Fields[2].AsFloat, Q.Fields[3].AsFloat, Q.Fields[4].AsFloat, Q.Fields[5].AsFloat, False )
        else CUpdateCumulsPointeMS( Q.Fields[0].AsString, Q.Fields[2].AsFloat, Q.Fields[3].AsFloat, 0, 0, False ) ;

    end;
    Ferme(Q1);
    Q.Next;
  end;
  Ferme(Q);
  FiniMove;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
function CTrouveContrePartie( vJournal : string ) : string ;
var lQuery : TQuery;
begin
  Result := '' ;
  try
    try
      lQuery := OpenSql('SELECT G_GENERAL FROM GENERAUX LEFT JOIN JOURNAL ON J_CONTREPARTIE=G_GENERAL' +
                        ' WHERE J_JOURNAL= "' + vJournal + '" ORDER BY G_GENERAL', True,-1,'',true);
      if not lQuery.Eof then
        Result := lQuery.FindField('G_GENERAL').AsString;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CTrouveContrePartie');
    end;

  finally
    Ferme(lQuery);
  end;
end;

{28/09/06 : gestion du Multi dossiers avec vDossier}
procedure RecalculTotPointeNew1( vCompteOuJournal : string; vDossier : string = '');
var
  lSt, lStCompteMAJ : string;
  //CodeDev : string;
  //Q1    : TQuery;
  Q       : TQuery;

  lTotDebPtp : Double;
  lTotCrePtp : Double;
  lTotDebPtd : Double;
  lTotCrePtd : Double;

begin
  // Recherche du compte de contrepartie car pointage sur journal
  if VH^.PointageJal then
  begin
    lStCompteMaj := CTrouveContrePartie( vCompteOuJournal );
    {JP 17/07/07 : FQ 20605 : ajout du test vCompteOuJournal <> '', car sinon, c'est normal que lStCompteMaj = ''}
    if (lStCompteMaj = '') and (vCompteOuJournal <> '') then
    begin
      PGIError('Impossible de trouver le compte de contrepartie du journal : ' + vCompteOuJournal, 'Traitement annulé' );
      Exit;
    end
  end
  else
    lStCompteMaj := vCompteOuJournal;

  if not EstTablePartagee( 'GENERAUX' ) then begin
    // Mise à zéro des totaux pointés du compte
    lSt := 'UPDATE GENERAUX SET G_TOTDEBPTP = 0, G_TOTCREPTP = 0, G_TOTDEBPTD = 0, G_TOTCREPTD = 0';
    if vCompteOuJournal <> '' then
    begin
      lSt := lSt + ' WHERE G_GENERAL = "' + lStCompteMAJ + '"';
    end;
    ExecuteSQL(lSt);
  end
  else
    {JP 28/09/06 : Gestion du Multi sociétés}
    CReinitCumulsPointeMS( lStCompteMAJ, vDossier) ;

  if VH^.PointageJal then
    lSt := 'SELECT E_JOURNAL, E_DEVISE,'
  else
    lSt := 'SELECT E_GENERAL, E_DEVISE,';

  {JP 28/09/06 : Gestion du Multi sociétés}
  lSt := lSt + ' SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT, ' +
          'SUM(E_DEBITDEV) DEBITDEV, SUM(E_CREDITDEV) CREDITDEV FROM ' + GetTableDossier(vDossier, 'ECRITURE') + ' ' +
          'LEFT JOIN GENERAUX on G_GENERAL = E_GENERAL ' +
          'WHERE E_REFPOINTAGE <> "" AND E_QUALIFPIECE = "N" AND (E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H" ) ';

  // Gestion du VH^.ExoV8
  if VH^.ExoV8.Code <> '' then
    lSt := lSt + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';

  if VH^.PointageJal then
  begin
    if vCompteOuJournal <> '' then
      lSt := lSt + ' AND E_JOURNAL = "' + vCompteOuJournal + '" AND E_GENERAL <> "' + lStCompteMaj + '"';

    lSt := lSt + ' GROUP BY E_JOURNAL, E_DEVISE ORDER BY E_JOURNAL, E_DEVISE';
  end
  else
  begin
    lSt := lSt + ' AND G_NATUREGENE = "BQE" AND G_POINTABLE = "X"';
    if vCompteOuJournal <> '' then
      lSt := lSt + ' AND E_GENERAL = "' + lStCompteMaj + '" ';

    lSt := lSt + ' GROUP BY E_GENERAL, E_DEVISE ORDER BY E_GENERAL, E_DEVISE';
  end;

  Q := OpenSQL(lSt, TRUE);
  //FetchSQLODBC(Q);
  InitMove(RecordsCount(Q), '');
  while not Q.Eof do
  begin
    MoveCur(FALSE);
    {JP 31/05/07 : FQ 20483 : ce code est inutile et dans sa logique (le recalcul des soldes part de
                   la table ECRITURE) et dans les faits, car juste au dessus on met le totaux à 0
    lSt := 'SELECT G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD, ' +
           'BQ_DEVISE FROM GENERAUX ' +
           'LEFT JOIN BANQUECP ON BQ_GENERAL = G_GENERAL' +
           ' AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"'; // 24/10/2006 YMO Multisociétés

    if VH^.PointageJal then
      lStCompteMaj := CTrouveContrePartie( vCompteOuJournal )
    else
      lStCompteMaj := vCompteOuJournal;

    lSt := lSt + ' WHERE G_GENERAL = "' + lStCompteMaj + '"';

    // Récupération des montants du compte à mettre à jour
    Q1 := OpenSQL(lSt, True);

    lTotDebPtp := Q.FindField('DEBIT').AsFloat + Q1.FindField('G_TOTDEBPTP').AsFloat;
    lTotCrePtp := Q.FindField('CREDIT').AsFloat + Q1.FindField('G_TOTCREPTP').AsFloat;

    // Devise Ecriture = Devise Compte
    if CodeDev = Q.FindField('E_DEVISE').AsString then
    begin
      lTotDebPtd := Q1.FindField('G_TOTDEBPTP').AsFloat;
      lTotCrePtd := Q1.FindField('G_TOTCREPTP').AsFloat;
    end
    else
    begin
      lTotDebPtd := Q.FindField('DEBITDEV').AsFloat  + Q1.FindField('G_TOTDEBPTP').AsFloat;
      lTotCrePtd := Q.FindField('CREDITDEV').AsFloat + Q1.FindField('G_TOTCREPTP').AsFloat;
    end;
    Ferme(Q1);
    }
    lTotDebPtp   := Q.FindField('DEBIT').AsFloat ;
    lTotCrePtp   := Q.FindField('CREDIT').AsFloat;
    lTotDebPtd   := Q.FindField('DEBITDEV').AsFloat ;
    lTotCrePtd   := Q.FindField('CREDITDEV').AsFloat;
    {JP 17/07/07 : Ajout du compte, car l'on se trouve dans une boucle}
    if vCompteOuJournal = '' then begin
      if VH^.PointageJal then begin
        lStCompteMaj := CTrouveContrePartie(Q.FindField('E_JOURNAL').AsString);
        if (lStCompteMaj = '') then begin
          PGIError('Impossible de trouver le compte de contrepartie du journal : ' + Q.FindField('E_JOURNAL').AsString, 'Traitement annulé' );
          Continue;
        end;
      end
      else
        lStCompteMAJ := Q.FindField('E_GENERAL').AsString;
    end;

    // Mise à jour des montants
    if not EstTablePartagee( 'GENERAUX' ) then begin
      {JP 17/07/07 : Gestion du compte général}
      lSt := 'UPDATE GENERAUX SET G_TOTDEBPTP = ' + VariantToSql(lTotDebPtp) + ',' +
                                 'G_TOTCREPTP = ' + VariantToSql(lTotCrePtp) + ',' +
                                 'G_TOTDEBPTD = ' + VariantToSql(lTotDebPtd) + ',' +
                                 'G_TOTCREPTD = ' + VariantToSql(lTotCrePtd) + ' ';
      if lStCompteMaj <> '' then lSt := lSt + 'WHERE G_GENERAL = "' + lStCompteMAJ + '"';
      ExecuteSQL(lSt);
    end
    else
      {MAJ table cumuls si besoin
       JP 28/09/06 : Gestion du multi sociétés}
      CUpdateCumulsPointeMS( lStCompteMAJ, lTotDebPtp, lTotCrePtp, lTotDebPtd, lTotCrePtd, True, vDossier ) ;

    Q.Next;
  end;
  Ferme(Q);
  FiniMove;
end;

procedure UpDateDecoupeEcr(SetSQL, WhereSQL: string; Exo: tExoDate);
var
  DMin, DMax, DD1, DD2: TDateTime;
  ListeJ: HTStrings;
  Q, QExo: TQuery;
  i, iper, Delta: integer;
  StSQL, StWhereExo: string;
  //    OkDate : Boolean ;
begin
  // Lecture des journaux
  ListeJ := HTStringList.Create;
  Q := OpenSQL('Select J_JOURNAL from JOURNAL', True,-1,'',true);
  while not Q.EOF do
  begin
    ListeJ.Add(Q.Fields[0].AsString);
    Q.Next;
  end;
  Ferme(Q);
  InitMove(1000, '');
  // Balayage des écritures avec découpe
  for i := 0 to ListeJ.Count - 1 do
  begin
    StWhereExo := '';
    if Exo.Code <> '' then
      StWhereExo := ' WHERE EX_EXERCICE="' + Exo.Code + '" ';
    QExo := OpenSQl('Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE'
      + StWhereExo, True,-1,'',true);
    while not QExo.EOF do
    begin
      DMin := QExo.Fields[1].AsDateTime;
      DMax := QExo.Fields[2].AsDateTime;
      Delta := Round((DMax - DMin) / 10);
      for iper := 1 to 10 do
      begin
        MoveCur(FALSE);
        if iper < 10 then
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DD1 + Delta;
        end
        else
        begin
          DD1 := DMin + (iper - 1) * Delta;
          DD2 := DMax;
        end;
        StSQL := 'UPDATE ECRITURE SET ' + SetSQL + ' WHERE E_JOURNAL="' +
          ListeJ[i] + '" AND E_EXERCICE="' + QExo.Fields[0].AsString
          + '" AND E_DATECOMPTABLE>="' + UsDateTime(DD1) +
            '" AND E_DATECOMPTABLE<"' + UsDateTime(DD2) + '"';
        if WhereSQL <> '' then
          StSQL := StSQL + ' ' + WhereSQL;
        BeginTrans;
        ExecuteSQL(StSQL);
        CommitTrans;
      end;
      QExo.Next;
    end;
    Ferme(QExo);
  end;
  ListeJ.Clear;
  ListeJ.Free;
  FiniMove;
end;

procedure InitCodeAccept(WhereSup: string; Exo: TExoDate);
var
  LMP: HtStringList;
  i: Integer;
  St, MPLu, Acc, CatLu, StSup: string;
begin
  ChargeMPACC;
  LMP := HtStringList.Create;
  for i := 0 to VH^.MPACC.Count - 1 do
  begin
    St := VH^.MPACC[i];
    MPLu := ReadtokenSt(St);
    Acc := ReadtokenSt(St);
    CatLu := ReadTokenSt(St);
    if CatLu = 'LCR' then
      LMP.Add(MpLU + ';' + Acc + ';');
  end;
  for i := 0 to LMP.Count - 1 do
  begin
    St := LMP[i];
    MPLu := ReadtokenSt(St);
    Acc := ReadtokenSt(St);
    StSup := 'AND E_AUXILIAIRE<>"" AND E_MODEPAIE="' + MPLu +
      '" AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") ' +
      'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") AND (E_NATUREPIECE="FC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OD")'
      + WhereSup;
    UPDATEDECOUPEECR(' E_CODEACCEPT="' + ACC + '" ', StSup, Exo);
  end;
  LMP.Free;
end;

procedure UpdateLibelleModeleCHQ(NatModele: string);
var
  Q: TQuery;
  St: string;
  TOBB, TOBL: TOB;
  CodeCh: string;
  Lib, Jal, Cpt: string;
  i: Integer;
  IdentModele: string;
begin
  if NatModele = 'LCH' then
    IdentModele := 'BQ_LETTRECHQ'
  else if NatModele = 'LVI' then
    IdentModele := 'BQ_LETTREVIR'
  else if NatModele = 'LPR' then
    IdentModele := 'BQ_LETTREPRELV'
  else
    Exit;
  St := 'SELECT ' + IdentModele + ',BQ_GENERAL FROM BANQUECP WHERE ' +
    IdentModele + '<>""';
  Q := OpenSQL(St, True,-1,'',true);
  TOBB := TOB.Create('', nil, -1);
  TOBB.LoadDetailDB('BANQUECP', '', '', Q, False, True);
  Ferme(Q);
  for i := 0 to TOBB.Detail.Count - 1 do
  begin
    TOBL := TOBB.Detail[i];
    CodeCH := TOBL.GetValue(IdentModele);
    Lib := '';
    Cpt := TOBL.GetValue('BQ_GENERAL');
    Q :=
      OpenSQL('SELECT MO_LIBELLE FROM MODELES WHERE MO_TYPE="L" AND MO_NATURE="' +
      NatModele + '" AND MO_CODE="' + CodeCH + '" ', TRUE,-1,'',true);
    if not Q.Eof then
      Lib := Q.Fields[0].AsString;
    Ferme(Q);
    if Lib <> '' then
    begin
      Jal := '';
      Q := OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="' + CPT +
        '" AND J_NATUREJAL="BQE"', TRUE,-1,'',true);
      if not Q.Eof then
        Jal := Q.Fields[0].AsString;
      Ferme(Q);
      if Pos(';', Lib) = 0 then
      begin
        if NatModele = 'LCH' then
        else if NatModele = 'LVI' then
          Jal := 'VIR'
        else if NatModele = 'LPR' then
          Jal := 'PRL'
        else
          Exit;
        Lib := Jal + ';' + V_PGI.DevisePivot + ';' + Lib;
        Lib := Copy(Lib, 1, 70);
        ExecuteSQL('UPDATE MODELES SET MO_LIBELLE="' + Lib +
          '" WHERE MO_TYPE="L" AND MO_NATURE="' + NatModele + '" AND MO_CODE="' +
          CodeCH + '" ');
      end;
    end;
  end;
  TOBB.Free;
end;


procedure UpdateLibelleToutModele;
begin
  UpdateLibelleModeleCHQ('LCH');
  UpdateLibelleModeleCHQ('LVI');
  UpdateLibelleModeleCHQ('LPR');
end;

end.

