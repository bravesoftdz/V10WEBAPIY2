{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Interface pour la gestion de la détaxe avec le logiciel Global
Suite ........ : Refund.
Mots clefs ... : FO
*****************************************************************}
unit FOGbRefund;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, FileCtrl,
  //OleCtrls, GRActiveX_TLB,
  HTB97, HCtrls, ParamSoc, UTOB, HEnt1, ExtCtrls;

type
  TFFOGbRefund = class(TForm)
    //    GbRefundChq: TGlobalRefundCheque;
    BExec: TToolbarButton97;
    TimerGR: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BExecClick(Sender: TObject);
    procedure TimerGRTimer(Sender: TObject);
  private
    { Déclarations privées }
    TOBPiece: TOB;
    TOBTiers: TOB;
    ErrorGR: integer;
    ReferenceGR: string;
  public
    { Déclarations publiques }
  end;

function FOLanceGlobalRefund(TOBPiece, TOBTiers: TOB; AvecFiche: boolean = False): boolean;

implementation

uses
  Ent1, EntGC, FactUtil, FOUtil, FactTob;

const
  NomFichierTrans = 'TXTRANS.TXT'; // nom du fichier de transaction de Global Refund
  NomFichierLog = 'TXLOG.TXT'; // nom du fichier de réponse de Global Refund

  {$R *.DFM}

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 25/02/2003
  Modifié le ... : 25/02/2003
  Description .. : Appel d'une fonction de GBREFUND.DLL
  Suite ........ :
  Suite ........ : Déclaration de la fonction en C++ :
  Suite ........ : void CALLBACK ExecGlobalRefund(long Operation, long
  Suite ........ : Service, const char *Reference, const char *Texte)
  Mots clefs ... :
  *****************************************************************}

procedure AppelDLLGlobalRefund;
type
  TExecGlobalRefund = procedure(Operation, Service: Longint; const Reference, Texte: PChar); stdcall;
  THandle = integer;
var
  Handle: THandle;
  ExecGBRefund: TExecGlobalRefund;
  Reference: PChar;
  Texte, DLLName: string;
  Operation, Service: Longint;
begin
  DLLName := Trim(GetParamSoc('SO_GCFODIREXECDETAXE'));
  if DLLName <> '' then DLLName := 'C:\CEGIDAPP'; // ????
  if DLLName <> '' then DLLName := IncludeTrailingBackslash(DLLName);
  DLLName := DLLName + 'GBREFUND.DLL';
  Handle := LoadLibrary(PChar(DLLName));
  if Handle <> 0 then
  begin
    @ExecGBRefund := GetProcAddress(Handle, 'ExecGlobalRefund');
    if @ExecGBRefund <> nil then
    begin
      Operation := 0;
      Service := 0;
      Reference := '';
      Texte := TraduireMemoire('Envoi de la pièce vers le logiciel de détaxe...');
      ExecGBRefund(Operation, Service, Reference, PChar(Texte));
    end;
    FreeLibrary(Handle);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Retourne le taux de TVA d'une ligne de document
Mots clefs ... : FO
*****************************************************************}

function NumeroCategorieTva: string;
var CodeTaxe: string;
  NumCat: integer;
begin
  CodeTaxe := GetParamSoc('SO_GCDEFCATTVA');
  if Copy(CodeTaxe, 1, 2) <> 'TX' then CodeTaxe := 'TX1';
  NumCat := StrToInt(Copy(CodeTaxe, 3, 1));
  if (NumCat < 1) or (NumCat > 5) then NumCat := 1;
  Result := IntToStr(NumCat);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Retourne le taux de TVA d'une ligne de document
Mots clefs ... : FO
*****************************************************************}

function TauxCategorieTva(TOBLigne: TOB): double;
var CodeTaxe, RegimeTaxe, FamilleTaxe: string;
  TOBT: TOB;
begin
  Result := 0;
  if TOBLigne = nil then Exit;

  CodeTaxe := GetParamSoc('SO_GCDEFCATTVA');
  FamilleTaxe := TOBLigne.GetValue('GL_FAMILLETAXE' + NumeroCategorieTva);
  RegimeTaxe := TOBLigne.GetValue('GL_REGIMETAXE');
  TOBT := VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF', 'TV_REGIME', 'TV_CODETAUX'],
    [CodeTaxe, RegimeTaxe, FamilleTaxe], False);
  if TOBT <> nil then Result := TOBT.GetValue('TV_TAUXVTE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Retourne le répertoire de stockage des fichiers Global
Suite ........ : Refund.
Mots clefs ... : FO
*****************************************************************}

function GRGetDirectory: string;
begin
  Result := Trim(GetParamSoc('SO_GCFODIRDETAXE'));
  if Result = '' then Result := 'C:\TEMP';
  if not DirectoryExists(Result) then ForceDirectories(Result);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Insère les enregistrements d'en-tête dans le fichier pour
Suite ........ : Global Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GREcritHeaderRecord(TOBPiece, TOBTiers: TOB; var FichierGR: TextFile);
var Enreg, TName, TCountry: string;
  CliOk: boolean;
begin
  if TOBPiece = nil then Exit;

  CliOk := ((TOBTiers <> nil) and (TOBPiece.GetValue('GP_TIERS') <> FOGetParamCaisse('GPK_TIERS')));
  if CliOk then
  begin
    TName := Trim(TOBTiers.GetValue('T_LIBELLE')) + ' ' + Trim(TOBTiers.GetValue('T_PRENOM'));
    TCountry := TOBTiers.GetValue('T_PAYS');
  end else
  begin
    TName := '';
    TCountry := GetParamSoc('SO_GCTIERSPAYS');
  end;

  Enreg := 'H1' // RecType  2c
  + FormatDateTime('yyyymmdd', TOBPiece.GetValue('GP_DATEPIECE')) // Date 8c AAAAMMJJ
  + StringOfChar(' ', 10) // NumRef 10c
  + StringOfChar(' ', 15) // Receipts 15c
  + '0' // Sentence 1c
  + Format('%-40.40s', [TName]) // TName 40 c
  + StringOfChar(' ', 20) // TCC := 20c
  + StringOfChar(' ', 4) // TCCDate 4c
  + StringOfChar(' ', 20) // TPassport 20c
  + Format('%-3.3s', [TCountry]) // TCountry 3c
  + '0' // VATOff 1c
  + StringOfChar(' ', 8) // DateArrival 8c
  + StringOfChar(' ', 8) // DateDepartment 8c
  + StringOfChar(' ', 8); // DateDelivering 8c
  Writeln(FichierGR, Enreg);

  if CliOk then
  begin
    TName := Trim(TOBTiers.GetValue('T_CODEPOSTAL')) + ' ' + Trim(TOBTiers.GetValue('T_VILLE'));

    Enreg := 'H2' // RecType  2c
    + Format('%-40.40s', [TOBTiers.GetValue('T_ADRESSE1')]) // TAddr1 40 c
    + Format('%-40.40s', [TOBTiers.GetValue('T_ADRESSE2')]) // TAddr2 40 c
    + Format('%-40.40s', [TName]); // TAddr3 40 c
    Writeln(FichierGR, Enreg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Insère les enregistrements lignes dans le fichier pour Global
Suite ........ : Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GREcritLineRecord(TOBPiece, TOBTva: TOB; var FichierGR: TextFile);
var Enreg: string;
  Ind: integer;
  Qte, Taux, Base, Mnt: double;
  TOBL, TOBT: TOB;
begin
  if TOBPiece = nil then Exit;
  for Ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' then Continue;
    if TOBL.GetValue('GL_TYPEARTICLE') <> 'MAR' then Continue;
    if TOBL.GetValue('GL_ARTICLE') = '' then Continue;
    if TOBL.GetValue('GL_QTEFACT') < 0 then Continue;

    Qte := TOBL.GetValue('GL_QTEFACT');
    Taux := TauxCategorieTva(TOBL);
    Base := TOBL.GetValue('GL_TOTALTTCDEV');
    Mnt := TOBL.GetValue('GL_TOTALTAXE' + NumeroCategorieTva);

    Enreg := 'RG' // RecType  2c
    + Format('%-40.40s', [TOBL.GetValue('GL_LIBELLE')]) // Goods 40 c
    + Format('%4.0f', [Qte]) // Qty 4 c
    + Format('%12.0f', [(Base * 100)]) // Price 12 c
    + Format('%4.0f', [(Taux * 100)]) // VAT rate 4 c
    + Format('%12.0f', [(Mnt * 100)]); // VAT amount 12 c
    Writeln(FichierGR, Enreg);

    TOBT := TOBTva.FindFirst(['TAUXTVA'], [Taux], False);
    if TOBT = nil then
    begin
      TOBT := TOB.Create('', TOBTva, -1);
      TOBT.AddChampSupValeur('TAUXTVA', Taux);
      TOBT.AddChampSupValeur('BASETVA', Base);
      TOBT.AddChampSupValeur('VALEURTVA', Mnt);
    end else
    begin
      Base := Base + TOBT.GetValue('BASETVA');
      TOBT.PutValue('BASETVA', Base);
      Mnt := Mnt + TOBT.GetValue('VALEURTVA');
      TOBT.PutValue('VALEURTVA', Mnt);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Insère les enregistrements taxes dans le fichier pour Global
Suite ........ : Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GREcritTaxeRecord(TOBTva: TOB; var FichierGR: TextFile);
var Enreg: string;
  Ind: integer;
  TOBT: TOB;
  Taux, Base, Mnt: double;
begin
  if TOBTva = nil then Exit;
  for Ind := 0 to TOBTva.Detail.Count - 1 do
  begin
    TOBT := TOBTva.Detail[Ind];
    Taux := TOBT.GetValue('BASETVA') * 100;
    Base := TOBT.GetValue('TAUXTVA') * 100;
    Mnt := TOBT.GetValue('VALEURTVA') * 100;

    Enreg := 'RT' // RecType  2c
    + Format('%12.0f', [Taux]) // Total 12 c
    + Format('%4.0f', [Base]) // VAT rate 4 c
    + Format('%12.0f', [Mnt]); // Total VAT 12 c
    Writeln(FichierGR, Enreg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Constitue à partir de la pièce le fichier pour Global Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GREcritFichier(TOBPiece, TOBTiers: TOB);
var NomFichier, NomDir: string;
  FichierGR: TextFile;
  TOBTva: TOB;
begin
  TOBTva := TOB.Create('', nil, -1);
  NomDir := GRGetDirectory;
  NomFichier := IncludeTrailingBackslash(NomDir) + NomFichierLog;
  DeleteFile(NomFichier);

  NomFichier := IncludeTrailingBackslash(NomDir) + NomFichierTrans;
  DeleteFile(NomFichier);

  AssignFile(FichierGR, NomFichier);
  if FileExists(NomFichier) then Reset(FichierGR) else Rewrite(FichierGR);
  try
    GREcritHeaderRecord(TOBPiece, TOBTiers, FichierGR);
    GREcritLineRecord(TOBPiece, TOBTva, FichierGR);
    GREcritTaxeRecord(TOBTva, FichierGR);
  finally
    CloseFile(FichierGR);
    TOBTva.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Interprête le fichier de réponse de Global Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GRLitReponse(var ErrorGR: integer; var ReferenceGR: string);
var NomFichier, NomDir: string;
  FichierLog: TextFile;
  Enreg: string;
begin
  NomDir := GRGetDirectory;
  NomFichier := IncludeTrailingBackslash(NomDir) + NomFichierLog;
  if not FileExists(NomFichier) then Exit;

  AssignFile(FichierLog, NomFichier);
  FileMode := 0;
  Reset(FichierLog);
  try
    Readln(FichierLog, Enreg);
    ErrorGR := StrToInt(Copy(Enreg, 1, 1));
    ReferenceGR := Trim(Copy(Enreg, 6, 10));
  finally
    CloseFile(FichierLog);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Met à jour la pièce avec les informations retournées par
Suite ........ : Global Refund.
Mots clefs ... : FO
*****************************************************************}

procedure GRMajPiece(ErrorGR: integer; ReferenceGR: string; TOBPiece: TOB);
var sSql: string;
  CleDoc: R_CleDoc;
begin
  if ErrorGR <> 0 then Exit;
  if TOBPiece = nil then Exit;

  if (TOBPiece.FieldExists('GP_DETAXE')) and
    (TOBPiece.FieldExists('GP_NUMDETAXE')) then
  begin
    TOBPiece.PutValue('GP_DETAXE', 'X');
    TOBPiece.PutValue('GP_NUMDETAXE', ReferenceGR);

    if V_PGI.IoError = oeOk then
    begin
      CleDoc := TOB2CleDoc(TOBPiece);
      sSql := 'UPDATE PIECE SET GP_DETAXE="X",GP_NUMDETAXE="' + ReferenceGR + '"'
        + ' WHERE ' + WherePiece(CleDoc, ttdPiece, False);
      if ExecuteSQL(sSql) <= 0 then V_PGI.IoError := oeUnknown;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/03/2003
Modifié le ... : 19/03/2003
Description .. : Lance l'interface pour la gestion de la détaxe avec le
Suite ........ : logiciel Global Refund.
Mots clefs ... : FO
*****************************************************************}

function FOLanceGlobalRefund(TOBPiece, TOBTiers: TOB; AvecFiche: boolean = False): boolean;
var XX: TFFOGbRefund;
  ErrorGR: integer;
  ReferenceGR: string;
begin
  if AvecFiche then
  begin
    XX := TFFOGbRefund.Create(Application);
    try
      XX.TOBPiece := TOBPiece;
      XX.TOBTiers := TOBTiers;
      XX.ShowModal;
    finally
      Result := (XX.ErrorGR = 0);
      XX.Free;
    end;
  end else
  begin
    SourisSablier;
    ErrorGR := -1;
    ReferenceGR := '';
    GREcritFichier(TOBPiece, TOBTiers);
    AppelDLLGlobalRefund;
    SourisSablier;
    GRLitReponse(ErrorGR, ReferenceGR);
    GRMajPiece(ErrorGR, ReferenceGR, TOBPiece);
    Result := (ErrorGR = 0);
  end;
  Application.ProcessMessages;
  SourisNormale;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : FormCreate
Mots clefs ... : FO
*****************************************************************}

procedure TFFOGbRefund.FormCreate(Sender: TObject);
begin
  ErrorGR := -1;
  ReferenceGR := '';
  TOBPiece := nil;
  TOBTiers := nil;
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : FormDestroy
Mots clefs ... : FO
*****************************************************************}

procedure TFFOGbRefund.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : BExecClick
Mots clefs ... : FO
*****************************************************************}

procedure TFFOGbRefund.BExecClick(Sender: TObject);
begin
  //GbRefundChq.ControlInterface ;
  //GbRefundChq.Exec (0, 0, '') ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 24/02/2003
Modifié le ... : 24/02/2003
Description .. : TimerGRTimer
Mots clefs ... : FO
*****************************************************************}

procedure TFFOGbRefund.TimerGRTimer(Sender: TObject);
begin
  if not TimerGR.Enabled then Exit;
  TimerGR.Enabled := False;

  SourisSablier;
  GREcritFichier(TOBPiece, TOBTiers);
  AppelDLLGlobalRefund;
  SourisSablier;
  GRLitReponse(ErrorGR, ReferenceGR);
  GRMajPiece(ErrorGR, ReferenceGR, TOBPiece);
  SourisNormale;
  Close;
end;

end.
