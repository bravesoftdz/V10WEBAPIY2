unit MailingWord;

interface

uses
   classes,
   UTob,
   SysUtils, hqry, AGLUtilOLE, hctrls;

type
//  WordVariant = OleVariant;
  TModeExport = (meFull, meLight, meNone);
  TTypeExport = (teWord, teText, teNone);

type
   TMailing = class
   public
      FTOB : TOB;
      FQ : THQuery;
      //
      constructor Create;
      destructor Destroy; override;
      //
      procedure LanceExportDataSource;
      procedure LanceFusionDocument;
      procedure LanceFusionExist;
      procedure LanceFusionNouveau;
      procedure LanceFusionPrint;
      procedure LanceFusionVerif;
   private
      FTypeExport : TTypeExport;
      FInitOk : boolean;
      FModeExport : TModeExport;
      FInput : string;
      FOutput : string;
      FDirMaquette : string;
      FDirResultat : string;
      FDataFileName : string;
      FTmpFiles : TStrings;
      FWinWord : OleVariant;
      FWinNew : boolean;
      FActiveDocument : OleVariant;
      FDocDataSource : OleVariant;
      Ffichier : TextFile;
      FFieldNames : TStrings;
      FFormuleQ : TFormuleQ;

      procedure FusionVers(versQuoi: Cardinal);
      procedure FusionVersNewDocument;
      procedure FusionVersPrinter;
      procedure ShowDefaultDocument;
      procedure SelectDocumentAFusionner;
      function CreateFileDataSource(bFullExport : boolean = false): boolean;
      procedure CreateDataSourceFromQuery;
      function GetTempFileDataSource(ext: string = '.txt'): string;
      procedure LinkDocToDataSource;
      procedure CreateDataSourceFromTOB(laTob: TOB; Entete: boolean = true);
      procedure Ecrit(st: string; bWithLF: boolean = false);
      procedure EcritFeedback;
      procedure EcritSautLigne;
      procedure FermeFichier;
      function OuvreFichier: boolean;
      procedure Init(Q : THQuery ; T : TOB ; Input : string = ''; Output : string = '' ; FieldCaptions : string = '');
      procedure LanceFusionTOB;
      function GetLibelleFusion(Num: integer; Champ: string): string;
   end;

procedure LanceFusionNew (Q : THQuery ; Input : string = ''; FieldCaptions : string = '')  ;
procedure LanceFusionOpen (Q : THQuery; Input : string = ''; Output : string = '' ; FieldCaptions : string = '' )  ;
procedure LanceFusionWord (Q : THQuery ; Vers : string = 'FILE'; Input : string = ''; Output : string = '' ; FieldCaptions : string = '')  ;
procedure LanceFusionVerif (Q : THQuery ; Input : string = ''; Output : string = '' ; FieldCaptions : string = '')  ;
procedure LanceExportDataSource (Q : THQuery ; Output : string = ''; FieldCaptions : string = '')  ;
procedure LanceFusionTob (T : tob ; Input : string = ''; Output : string = ''; FieldCaptions : string = '')  ;

procedure LancePublipostage(Vers : string ;Input : string = ''; Output : string = ''; T : TOB = nil ; FieldCaptions : string = '')  ;
procedure AglPubliPostage (parms: array of variant; nb: integer )  ;
function AglIsWordDisponible (parms: array of variant; nb: integer ) : variant;

implementation

uses forms,dialogs,hent1,hmsgbox,hstatus,
{$IFDEF EAGLCLIENT}
     eFiche,eMul, eFichList,
{$ELSE}
     db,Fiche,Mul, FichList,
{$ENDIF}

     {$IFDEF VER150}
     Variants,
     {$ENDIF}

     M3FP;

var FMailing : TMailing;

procedure TMailing.Init(Q : THQuery ; T : TOB ; Input : string = ''; Output : string = '' ;
  FieldCaptions : string = '' );
var st : string;
begin
   CheckWord(FWinWord,FWinNew,false);
   if not OpenWord(false,FWinWord,FWinNew,false) then
      Exit;
   if FWinWord.Application.Visible then
      FWinWord.WindowState := wdWindowStateMinimize;
      
   FWinWord.Application.Visible := false;
   FQ := Q;
   FTOB := T;

   if Input <> '' then
   begin
      FInput := ExtractFileName( Input );
      FDirMaquette := ExtractFileDir( Input );
   end
   else
      FDirMaquette := ExtractFileDir( TempFileName );

   if Output <> '' then
   begin
      FOutput := ExtractFileName( Output );
      FDirResultat := ExtractFileDir( Output );
   end
   else
      FDirResultat := ExtractFileDir( TempFileName );

st := FieldCaptions;
FFieldNames.Clear;
while st <> '' do FFieldNames.Add(ReadTokenSt(st));
end;

procedure LanceFusionTob (T : tob ; Input : string = ''; Output : string = '' ; FieldCaptions : string = '')  ;
begin
try
  FMailing.Init(nil,T,Input,Output,FieldCaptions);
  FMailing.LanceFusionTOB;
finally
//  Application.BringToFront;
end;
end;

procedure LanceExportDataSource (Q : THQuery ; Output : string = ''; FieldCaptions : string = '')  ;
begin
try
  FMailing.Init(Q, nil, '', Output, FieldCaptions );
  FMailing.LanceExportDataSource;
finally
  Application.Restore;
  Application.BringToFront;
end;
end;

procedure LanceFusionWord (Q : THQuery ; Vers : string = 'FILE'; Input : string = ''; Output : string = '';
   FieldCaptions : string = '')  ;
begin
   if not OfficeWinWordDispo then
      exit; // test si Word dispo

   try
      FMailing.Init(Q, nil, Input, Output, FieldCaptions );
      if Vers='PRINTER' then
         FMailing.LanceFusionPrint
      else
         FMailing.LanceFusionDocument;
   finally
      if Vers='PRINTER' then
       begin
        Application.Restore;       
        Application.BringToFront;
       end;
   end;
end;

procedure LanceFusionNew (Q : THQuery ; Input : string = ''; FieldCaptions : string = '')  ;
begin
   if not OfficeWinWordDispo then
      exit; // test si Word dispo
   try
      FMailing.Init(Q, nil, Input, '', FieldCaptions );
      FMailing.LanceFusionNouveau;
   finally
   //  Application.BringToFront;
   end;
end;


procedure LanceFusionOpen (Q : THQuery; Input : string = ''; Output : string = '' ;
   FieldCaptions : string = '' )  ;
begin
   if not OfficeWinWordDispo then
      exit; // test si Word dispo

   try
      //*** Tests BM ***///
     FMailing.Init(Q, nil, Input, '', FieldCaptions );
     FMailing.LanceFusionExist;
   finally
      //  Application.BringToFront;
   end;
end;

procedure LanceFusionVerif (Q : THQuery ; Input : string = ''; Output : string = '' ; FieldCaptions : string = '')  ;
begin
if not OfficeWinWordDispo then exit; // test si Word dispo
try
  FMailing.Init(Q,nil,Input,Output,FieldCaptions);
  FMailing.LanceFusionVerif;
finally
//  Application.BringToFront;
end;
end;

procedure TMailing.ShowDefaultDocument;
begin
   if (VarType(FWinWord) = varEmpty) then
      exit;
   if (VarType(FActiveDocument) = varEmpty) then
      exit;
//   if FWinVer >= 9 then
      FActiveDocument.ActiveWindow.View.Type := wdPrintView;
//   else
//      FActiveDocument.ActiveWindow.View.Type := wdOnLineView;
      
   FWinWord.Application.Visible := true;
   FWinWord.WindowState := wdWindowStateMaximize;
   FWinWord.Activate;
end;

procedure TMailing.LinkDocToDataSource;
begin
   try
   begin
      FWinWord.Options.CheckGrammarAsYouType := False;
      FWinWord.Options.Pagination := False;

      if FInput = '' then
         FActiveDocument := FWinWord.Documents.Add
      else
         FActiveDocument := FWinWord.Documents.Open( FileName := FDirMaquette + '\' + FInput,
                                                     Revert := False ,
                                                     Format:=wdOpenFormatAuto);

      FActiveDocument.MailMerge.MainDocumentType := wdFormLetters;

      FActiveDocument.MailMerge.OpenDataSource ( Name:= FDataFileName ,
                                                Format := wdOpenFormatAuto,
                                                ConfirmConversions:=false ,
                                                ReadOnly:=False,
                                                LinkToSource:=false,
                                                AddToRecentFiles:=False);
   end
   finally
      FActiveDocument.MailMerge.EditMainDocument;
   end;
end;

procedure TMailing.FusionVers(versQuoi : Cardinal);
begin
   if (VarType(FWinWord) = varEmpty) then
      exit;
   try
   begin
      FActiveDocument.MailMerge.Destination := versQuoi;
{
  FActiveDocument.MailMerge.MailAsAttachment := False;
  FActiveDocument.MailMerge.MailAddressFieldName := '';
  FActiveDocument.MailMerge.MailSubject := '';
  FActiveDocument.MailMerge.SuppressBlankLines := True;
  FActiveDocument.MailMerge.DataSource.FirstRecord := wdDefaultFirstRecord;
  FActiveDocument.MailMerge.DataSource.LastRecord := wdDefaultLastRecord;
}
      FActiveDocument.MailMerge.Execute (False);
   end
   finally
      if (versQuoi=wdSendToNewDocument) and ( FOutPut <> '' ) then
      begin
         FWinWord.ActiveDocument.SaveAs( FDirResultat + '\' + FOutput );
   end;
      FWinWord.Application.Visible := (versQuoi=wdSendToNewDocument);
   end;
end;

procedure TMailing.FusionVersNewDocument;
begin
FusionVers(wdSendToNewDocument);
end;

procedure TMailing.FusionVersPrinter;
begin
FusionVers(wdSendToPrinter);
end;

constructor TMailing.Create;
begin
FInitOk := OfficeWinWordDispo;
FTmpFiles := TStringList.Create;
FWinNew:= true;
FWinWord := unassigned;
FActiveDocument := unassigned;
FInput := '';
FOutput := '';
FDirMaquette := '';
FDirResultat := '';
FDataFileName := '';
FTOB := nil;
FQ := nil;
FTypeExport := teNone;
FModeExport := meNone;
FFieldNames := TStringList.Create;
FFormuleQ := TFOrmuleQ.Create;
end;

destructor TMailing.Destroy;
var i : integer;
begin
if (IsWordLaunched) and (varType(FWinWord) <> varEmpty) then
  begin
  try
    FWinWord.Quit(wdDoNotSaveChanges);
  finally
    FWinWord := unassigned;
  end;
  end
  else FWinWord := unassigned;
if FTmpFiles<>nil then
  begin
  for I:=0 to FTmpFiles.Count-1 do
    if not DeleteFile(FTmpFiles[i]) then PGIInfo('Impossible de supprimer le fichier '+FTmpFiles[i],'Mailing');
  FTmpFiles.Free;
  end;
if FFieldNames<>nil then FFieldNames.Free;
FreeAndNil(FFormuleQ);
end;

//CREATION D'UN NOUVEAU DOCUMENT pour MAILING
//DATASOURCE : SCHEMA + JEU D'ESSAI
procedure TMailing.LanceFusionNouveau;
//var
//   sFicTmp_l : string;
begin
   //Export schema + sample data
   FDataFileName := GetTempFileDataSource;

   if not CreateFileDataSource then
      exit;
   //new document
   FInput := '';
   LinkDocToDataSource;
   ShowDefaultDocument;
end;

procedure TMailing.SelectDocumentAFusionner;
var OD : TOpenDialog;
begin
   //Open document a fusionner
   OD := TOpenDialog.Create(Application);
   try
   begin
      OD.Title := 'Sélectionner le document maquette à fusionner';
      OD.Filter := 'Documents Word (*.doc)|*.doc';
      OD.InitialDir := FDirMaquette;
      OD.FileName := FInput;

      if OD.execute then
      begin
         FInput := ExtractFileName( OD.FileName );
         FDirMaquette := ExtractFileDir( OD.FileName );
      end;
   end;
   finally
      OD.Free;
   end;
end;

function TMailing.CreateFileDataSource (bFullExport : boolean = false): boolean;
begin
   result := false;
   if bFullExport then
      FModeExport := meFull
   else
      FModeExport := meLight;
   if ExtractFileExt( FDataFileName ) = '.doc' then
      FTypeExport := teWord
   else if ExtractFileExt( FDataFileName ) = '.txt' then
      FTypeExport := teText
   else
      exit;
   if FTOB<>nil then
      CreateDataSourceFromTOB(FTob)
   else
   if FQ<>nil then
      CreateDataSourceFromQuery;
   if FDataFileName<>'' then
      FTmpFiles.Add(FDataFileName);
   result := FDataFileName<>'';
end;

//VERFICATION DE LA FUSION DANS UN NOUVEAU DOCUMENT
//DATASOURCE : SCHEMA + TOB DATA
procedure TMailing.LanceFusionTOB;
begin
FDataFileName:=GetTempFileDataSource ;
if not CreateFileDataSource(true) then exit;
//if FInput = '' then SelectDocumentAFusionner;
//if FInput<>'' then
//  begin
  LinkDocToDataSource;
  if FInput <> '' then
    begin
    if FOutput<>'' then FusionVersNewDocument;
    end;
  ShowDefaultDocument ;
//  end;
end;

//OUVERTURE D'UN DOCUMENT EXISTANT pour MAILING
//DATASOURCE : SCHEMA + JEU D'ESSAI
procedure TMailing.LanceFusionExist;
begin
   //Export schema + sample data
   FDataFileName:=GetTempFileDataSource ;
   if not CreateFileDataSource then
      exit;
   if FInput = '' then
      SelectDocumentAFusionner;
   if FInput <> '' then
   begin
      LinkDocToDataSource;
      ShowDefaultDocument;
   end;
end;

//LANCEMENT DE LA FUSION DANS UN NOUVEAU DOCUMENT
//DATASOURCE : SCHEMA + ALL DATA
procedure TMailing.LanceFusionDocument;
begin
   //Export schema + all data
   FDataFileName:=GetTempFileDataSource;

   if not CreateFileDataSource(true) then
      exit;
   // fusion vers document
   if FInput = '' then
      SelectDocumentAFusionner;
   if FInput<>'' then
   begin
      LinkDocToDataSource;
      FusionVersNewDocument;
   end;
end;

//VERFICATION DE LA FUSION DANS UN NOUVEAU DOCUMENT
//DATASOURCE : SCHEMA + ALL DATA
procedure TMailing.LanceFusionVerif;
begin
   //Export schema + all data
   FDataFileName:=GetTempFileDataSource ;

   if not CreateFileDataSource(true) then
      exit;
   if FInput = '' then
      SelectDocumentAFusionner;
   if FInput<>'' then
   begin
      LinkDocToDataSource;
      FusionVersNewDocument;
      ShowDefaultDocument ;
   end;
end;

//VERFICATION DE LA FUSION DANS L'IMPRIMANTE
//DATASOURCE : SCHEMA + ALL DATA
procedure TMailing.LanceFusionPrint;
begin
   //Export schema + all data
   FDataFileName:=GetTempFileDataSource;

   if not CreateFileDataSource(true) then
      exit;
   // fusion vers l'imprimante
   if FInput = '' then
      SelectDocumentAFusionner;
   if FInput<>'' then
   begin
      LinkDocToDataSource;
      FusionVersPrinter;
   end;
end;

//EXPORT DATASOURCE : SCHEMA + ALL DATA ou JEU D'ESSAI
procedure TMailing.LanceExportDataSource;
var SD : TSaveDialog;
begin
   SD := TSaveDialog.Create(Application);
   try
      SD.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Word (*.doc)|*.doc';
      SD.Title := 'Sélectionner le nom du fichier de données à créer';
      SD.DefaultExt := 'txt';
      if FDirResultat <> '' then
         SD.InitialDir := FDirResultat;
      if FOutput <> '' then
         SD.FileName := FOutput;

      //Export schema + all data
      if SD.execute then
      begin
         FDataFileName := SD.FileName;
         CreateFileDataSource(true);
      end;
   finally
      SD.Free;
   end;
end;

function TMailing.OuvreFichier : boolean;
begin
result := false;
if FTypeExport=teWord then
  begin
  FDocDataSource := FWinWord.Documents.Add;
  result := true;
  end
else if FTypeExport=teText then
  begin
//  {$I-} AssignFile(Ffichier, FDataFileName ); Rewrite(Ffichier); {$I+}
  AssignFile(Ffichier, FDataFileName ); {$I-} Rewrite(Ffichier); {$I+}
  if IoResult <> 0 then
    begin
    PGIInfo('Impossible d''ouvrir le fichier '+FDataFileName,'Erreur');
    result := false; exit;
    end else result := true;
  end;
end;

procedure TMailing.FermeFichier;
begin
if FTypeExport=teWord then
  begin
  FDocDataSource.SaveAs(FDataFileName);
  FDocDataSource.Close;
  end
else if FTypeExport=teText then CloseFile(Ffichier);
end;

procedure TMailing.Ecrit (st : string ; bWithLF : boolean = false);
begin
if FTypeExport=teWord then FWinWord.Selection.TypeText(st)
else if FTypeExport=teText then Write(Ffichier,St);
end;

procedure TMailing.EcritSautLigne ;
begin
if FTypeExport=teWord then
  begin
  FWinWord.Selection.TypeBackSpace;
  FWinWord.Selection.TypeParagraph;
  end
else if FTypeExport=teText then
  begin
  WriteLn(Ffichier,'');
  end;
end;

procedure TMailing.EcritFeedback ;
begin
if FTypeExport=teWord then
  FWinWord.Selection.TypeBackSpace; // efface ligne surnuméraire
end;

function TMailing.GetLibelleFusion (Num : integer ; Champ : string) : string;
var
   st : string;
begin
   st := '??';
   if FFieldNames.Count <> 0 then
   begin
      if Num < FFieldNames.Count then
         st := FFieldNames[Num];
   end
   else
   begin
      St := ChamptoLibelle(Champ);
      if st = '??' then
         st := Champ;
   end;
   if St = '??' then
      st := 'Champ n°' + IntToStr(Num);
   result := '"'+st+'"';
end;

procedure TMailing.CreateDataSourceFromTOB (laTob : TOB ; Entete : boolean = true);
var i,ii : integer;
    sep,st,stVal : string;
    bSautLigne : boolean;
begin
sep := '|';
bSautLigne := false;
if (laTob=nil) or (FDataFileName='') then exit;
if Entete then
    BEGIN
    if OuvreFichier then else exit;
    ii := 0;//    i := 1;
    For i:=1 to laTob.NbChamps do
      begin
      st := GetLibelleFusion (i-1,V_PGI.DEChamps[laTob.NumTable,i].Nom);
      bSautLigne := true;
      Ecrit(st+sep);
      ii := i;
      end;
    For i:=0 to laTob.ChampsSup.Count-1 do
      begin
      st := GetLibelleFusion (ii+i,TCS(laTob.ChampsSup[i]).Nom);
      bSautLigne := true;
      Ecrit(st+sep);
      end;
      if bSautLigne then EcritSautLigne;
    END ;
bSautLigne := false;
// pas de la 1ere mere si detail
if laTob.Detail.Count=0 then
For i:=1 to laTob.NbChamps do
  BEGIN
  Case VarType(laTob.Valeurs[i]) of
      varNull :   stVal := '';
      varOleStr:  stVal := '';
      varDate :   stVal := DateTimeToStr(VarAsType(laTob.Valeurs[i], vardate));
      varDouble : stVal := FloatToStr(VarAsType(laTob.Valeurs[i], vardouble));
      varSingle : stVal := FloatToStr(VarAsType(laTob.Valeurs[i], varsingle));
      varInteger :stVal := IntToStr(VarAsType(laTob.Valeurs[i], varInteger));
      else        stVal := VarAsType(laTob.Valeurs[i], varString);
    END ;
    bSautLigne := true;
    Ecrit(stVal+sep);
  END ;
// pas de la 1ere mere si detail
if laTob.Detail.Count=0 then
For i:=1000 to laTob.ChampsSup.Count+999 do
  BEGIN
  Case VarType(laTob.Valeurs[i]) of
    varNull :   stVal := '' ;
    varOleStr:  stVal := '';
    varDate :   stVal := DateTimeToStr(VarAsType(laTob.Valeurs[i], vardate));
    varDouble : stVal := FloatToStr(VarAsType(laTob.Valeurs[i], vardouble));
    varSingle : stVal := FloatToStr(VarAsType(laTob.Valeurs[i], varsingle));
    varInteger :stVal := IntToStr(VarAsType(laTob.Valeurs[i], varInteger));
    else        stVal := VarAsType(laTob.Valeurs[i], varString);
    END ;
    bSautLigne := true;
    Ecrit(stVal+sep);
  END ;
if bSautLigne then EcritSautLigne;
for i:=0 to laTob.Detail.Count-1 do
  CreateDataSourceFromTOB(laTob.Detail[i],False) ; // laTob.Parent=nil : cas d'un detail
//
if bSautLigne then EcritFeedback;
if Entete then FermeFichier;
end;

procedure TMailing.CreateDataSourceFromQuery;
var C,LeMax,R : longint ;
    stval,sep,St : String ;
    okok : boolean;
begin
if (FQ=nil) or (FDataFileName='') then exit;
if OuvreFichier then else exit;
sep := '|';
with FQ do
  begin
  DisableControls ;
  First;
  Lemax:=1000 ;
  initmove(LeMax,'') ;
{$IFNDEF EAGLCLIENT}
  FFormuleQ.InitQ(FQ,FQ.Criteres);
  For C:=0 to FieldCount-1 do
    begin
    st := GetLibelleFusion(C,Fields[C].FieldName);
    Ecrit (st+sep) ;
    end;
{$ELSE}
  FFormuleQ.InitQ(FQ.TQ,FQ.Criteres);
  For C:=0 to TQ.FieldCount-1 do
    begin
    st := GetLibelleFusion(C,TQ.Fields[C].FieldName);
    Ecrit (st+sep) ;
    end;
{$ENDIF}
  FFormuleQ.Assign(FQ.FormuleQ);
  EcritSautLigne;
  R :=0;
  While Not EOF do
     BEGIN
     if (FModeExport=meLight) and (R=5) then break; // schema + qques lignes d'exemples
     if V_PGI.CergEF then okok:=(FindField('PT_ENTETE')<>Nil) AND
                                (FindField('PT_BOLD').AsString<>'X') AND
                                (FindField('PT_ENTETE').AsString<>'X')
        else okok:=TRUE ;
     If okok then
        BEGIN
        if MoveCur(True) then ;
{$IFNDEF EAGLCLIENT}
        for C:=0 to FieldCount-1 do
          BEGIN
          stVal:='';
          if (Fields[C].DataType in [ftDate,ftTime,ftDateTime]) and (Fields[C].AsDateTime<=iDate1900) then
          else if Fields[C].AsString = W_W then stVal := FFormuleQ.Compute(Fields[C].FieldName)
                                           else stVal := Fields[C].AsString;
          Ecrit (stVal+sep ) ;
          END ;
{$ELSE}
        for C:=0 to TQ.FieldCount-1 do
          BEGIN
          stVal:='';
          if (TQ.Fields[C].DataType in [ftDate,ftTime,ftDateTime]) and (TQ.Fields[C].AsDateTime<=iDate1900) then
          else if TQ.Fields[C].AsString = W_W then stVal := FFormuleQ.Compute(TQ.Fields[C].FieldName)
                                              else stVal := TQ.Fields[C].AsString;
          Ecrit (stVal+sep ) ;
          END ;
{$ENDIF}
        EcritSautLigne;
        END ;
     Next ;
     inc(R);
     END ;
  First ;
  EnableControls ;
  END ;
{$IFNDEF EAGLCLIENT}
{$ELSE}
{$ENDIF}
FiniMove ;
EcritFeedback;
FermeFichier;
end;

function TMailing.GetTempFileDataSource (ext : string = '.txt') : string;
var
   fic : string;
begin
   fic := TempFileName;
   ext:='.doc';// pb sur txt....
   if ext='.txt' then
      FTypeExport := teText
   else
   if ext='.doc' then
      FTypeExport := teWord
   else
      FTypeExport := teNone;
   result := ChangeFileExt(fic,ext);
   FTmpFiles.Add(fic); // pour le supprimer
end;

////////////////////////////
//GESTION DES MAILING WORD//
////////////////////////////
Function FindFormQuery(F : TForm) : TDataset ;
Var DS : TDataset ;
BEGIN
  DS:=NIL ;
{$IFDEF EAGLCLIENT}
  if F is TFFiche then DS:=TDataset(TFFiche(F).QFiche) ;
  if F is TFFicheListe then DS:=TDataset(TFFicheListe(F).Ta) ;
  if F is TFMul then DS:=TDataset(TFMul(F).Q.TQ) ;
{$ELSE}
  if F is TFFiche then DS:=TDataset(F.FindComponent('QFiche')) ;
  if F is TFFicheListe then DS:=TDataset(F.FindComponent('Ta')) ;
  if F is TFMul then DS:=TDataset(F.FindComponent('Q')) ;
{$ENDIF}
Result:=DS ;
END ;

procedure LancePublipostage(Vers : string ;Input : string = ''; Output : string = ''; T : TOB = nil ; FieldCaptions : string = '');
begin
  if ((Vers='TOB') or (Vers='MONOFICHIER')) then
    LanceFusionTOB(T,input,output,FieldCaptions);
end;

procedure AglPubliPostage (parms: array of variant; nb: integer )  ;
var
   vers,input,output,Names : string;
    F: TForm ;
    DS : TDataSet ;
    T : TOB;
    Q : THQuery;
begin
  F := TForm  (LongInt (parms[0])) ;
  DS:=FindFormQuery(F) ;
  Q := THQuery(DS);
  Vers := Trim(uppercase(string (parms[1])));
  input := Trim(string (parms[2]));
  output := Trim(string (parms[3]));
  T := TOB (LongInt (parms[4])) ;
  Names := Trim(string (parms[5]));
  if ((Vers='NEW') or (Vers='NOUVEAU')) then LanceFusionNew(Q,Names) else
  if ((Vers='OPEN') or (Vers='OUVRIR')) then LanceFusionOpen(Q,input,output,Names) else
  if ((Vers='FILE') or (Vers='FICHIER')) then LanceFusionWord(Q,Vers,input,output,Names) else
  if ((Vers='PRINTER') or (Vers='IMPRIMANTE')) then LanceFusionWord(Q,Vers,input,Names) else
  if ((Vers='EXPORT') or (Vers='DONNEES')) then LanceExportDataSource (Q,Names) else
  if ((Vers='VERIF') or (Vers='VERIFICATION')) then LanceFusionVerif(Q,input,output,Names) else
  if ((Vers='TOB') or (Vers='MONOFICHIER')) then LanceFusionTOB(T,input,output,Names);
end;

function AglIsWordDisponible (parms: array of variant; nb: integer ) : variant;
begin
result := (FMailing<>nil) and (FMailing.FInitOk);
end;

procedure initializeMailing();
begin
  RegisterAglFunc ( 'IsWordDisponible', True, 0, AglIsWordDisponible) ;
  RegisterAglProc ( 'PubliPostage', True, 5, AglPubliPostage) ;
  FMailing := TMailing.Create;
end;

procedure finalizeMailing;
begin
if FMailing<>nil then FMailing.Free;
end;

//***  Tests BM ***//
function GetFInput : string;
begin
   result := FMailing.FInput;
end;
//*** Tests BM ***///  
function GetFOutput : string;
begin
   result := FMailing.FOutput;
end;
//*** Tests BM ***///
function GetFDataFileName : string;
begin
   result := FMailing.FDataFileName;
end;


Initialization
initializeMailing();
finalization
finalizeMailing;

end.
