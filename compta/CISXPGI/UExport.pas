{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 06/11/2002
Modifié le ... : 06/11/2002
Description .. : Unité export du fichier TRA
Mots clefs ... :
*****************************************************************}

unit UExport;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms, Dialogs,
  Vierge, Grids, Hctrls, StdCtrls, Mask, ComCtrls, HSysMenu
  ,HPanel, UiUtil, db, 
{$IFDEF CISXPGI}
  ExtCtrls,
{$ENDIF}
   uDbxDataSet, Variants, ADODB,  HTB97,
   inifiles
  ,UTOB,UTOZ
  ,HmsgBox, Hent1
  ,UScript
  , UControlParam
  , ed_tools
{$IFNDEF EAGLCLIENT}
  ,Fe_Main
  ,EdtREtat {LanceEtatTob}
{$ELSE}
  ,MainEagl
  ,UScriptTob
  ,UtileAgl {LanceEtatTob}
{$ENDIF}
{$IFDEF CISXPGI}
  ,ULibCpContexte
{$ENDIF}
  ,HStatus;
type
  TFUExport = class(TFVierge)
    PageControl1: TPageControl;
    GD: THGrid;
    HPanel1: THPanel;
    LTable: TLabel;
    Table: THCritMaskEdit;
    TFListefichiers: TLabel;
    FListefichiers: THCritMaskEdit;
    BValide: TToolbarButton97;
    LFichier: TLabel;
    FILENAME: THCritMaskEdit;
    Memefichier: TCheckBox;
    Label12: TLabel;
    Profile: THCritMaskEdit;
    BCompress: TCheckBox;
    procedure FListefichiersElipsisClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure TableElipsisClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure GDDblClick(Sender: TObject);
    procedure GDSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure bDefaireClick(Sender: TObject);
    procedure TableExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TableClick(Sender: TObject);
    procedure ProfileElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
    CurrentDir    : string;
    TobExport,TF  : TOB;
    Ascript       : TScript;
    Domaine       : string;
    SavRow,SavCol : Integer;
    FSaisieVar    : TFVierge;
    ActionClose   : Boolean;
    FicheSortie   : string;
    DatedebutExo  : TDateTime;
    DatefinExo    : TDateTime;
    TOBVar        : TOB;
    OkCharge      : Boolean;
    ScriptSQL     : Boolean;
    Commande      : string;
    OkParamcount  : Boolean;
    ModeRlv,LPays : string;
    procedure InitScript(table : string; OkC:Boolean=FALSE);
    procedure  InitDelim(Codedelim : string);
    procedure FileClickzip (Filen : string);
    function  SaisieVariable (fich : string; TV : TOB): Boolean;
    procedure BValideClickVar(Sender: TObject);
  public
{$IFNDEF CISXPGI}
    { Déclarations publiques }
    procedure savetofile(AScriptName : string);
{$ENDIF}
  end;

var
  FUExport: TFUExport;

Procedure ExportDonneesCisx(PP: THPanel; Cmd : string=''; Mode : string='EXPORT'; Lepays : string='');
procedure ExecuteExportFichier (var F : TextFile; var TOBExp, TEcr, TReleve : Tob; FScript : TScript; ModeRlV : string; Rlv, Etebac : Boolean; TSV : TSVControle;  var DatedebutExo, DatefinExo : TDateTime; Lepays : string);

{$IFDEF CISXPGI}
function LanceExportCisx (Cmd : string; Mode : string='EXPORT'; Lepays : string='') : Boolean;
{$ENDIF}
implementation

uses
{$IFNDEF CISXPGI}
UMulScript,
USelectCorresp,
{$ENDIF}
UDMIMP
,UExecute
{$IFNDEF EAGLCLIENT}
,UVIEW
{$ENDIF}
,UDefVar
,UAssistConcept
,UScriptDelim
;

{$R *.DFM}

Procedure ExportDonneesCisx(PP: THPanel; Cmd : string=''; Mode : string='EXPORT'; Lepays : string='') ;
BEGIN
     FUExport :=TFUExport.Create(Application) ;
     with FUExport do
     begin
          Commande := Cmd;
          ModeRlv  := Mode;
          LPays    := Lepays;
          if Cmd <> '' then
             OkParamCount := TRUE
          else
             OkParamcount := ParLigneCommande; // OkParamcount remplace  (ParamCount > 0)
          if (PP = Nil) and (Cmd = '') then
          BEGIN
                Try
                   FUExport.ShowModal ;
                Finally
                       FUExport.Free;
          End;
          END else
          BEGIN
               InitInside(FUExport, PP);
               FUExport.Show;
{$IFDEF CISXPGI}
               if OkParamcount  then
                FUExport.close;
{$ENDIF}
          END ;
     end;
END ;

{$IFDEF CISXPGI}
function LanceExportCisx (Cmd : string; Mode : string='EXPORT'; Lepays : string='') : Boolean;
var
i        : integer;
begin
   TCPContexte.GetCurrent.VarCisx.CHARGECIX(Cmd);
   With GetInfoVHCX do
   begin
       if (ATOB <> nil) and (ATob.detail.count <> 0) then
       begin
             For i:= 0 to ATOB.detail.count-1 do
             begin
                Directory                      := ATOB.detail[i].getvalue('REPERTOIRE');
                Script                         := ATOB.detail[i].getvalue('SCRIPT');
                NomFichier                     := ATOB.detail[i].getvalue('NOMFICHIER');
                ListeFichier                   := ATOB.detail[i].getvalue ('LISTEFICHIER');
                Domaine                        := ATOB.detail[i].getvalue ('Domaine');
                Monotraitement                 := ATOB.detail[i].getvalue ('MONOTRAITEMENT');
                compress                       := ATOB.detail[i].getvalue ('COMPRESS');
                ExportDonneesCisx(nil, Cmd, Mode, Lepays);
             end;
       end
       else
            ExportDonneesCisx(nil, Cmd, Mode, Lepays);
   end;
    TCPContexte.GetCurrent.Release;
   Result := TRUE;
end;
{$ENDIF}

procedure TFUExport.FListefichiersElipsisClick(Sender: TObject);
var
  I   : integer;
  NN  : string;
begin
    if (FListefichiers.Text <> '') then
    begin
         CurrentDir := ExtractFileDir (FListefichiers.Text);
         SetCurrentDirectory(PChar(CurrentDir));
    end;
    FListefichiers.Text := '';
    with Topendialog.create(Self) do
    begin
         InitialDir := CurrentDir;
         Title := ' Choisir la liste des fichiers à traiter';
         Options := [ofFileMustExist, ofHideReadOnly, ofAllowMultiSelect ];
         Filter := 'Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*';
         FilterIndex := 2;
         if Execute then
         begin
         with Files do
              for I := 0 to Count - 1 do
              begin
              if I = 0 then FListefichiers.Text := Strings[I]
              else FListefichiers.Text := FListefichiers.Text+';'+Strings[I];
              end;
         end;
         NN := FileName;
         CurrentDir := ExtractFileDir (NN);
         SetCurrentDirectory(PChar(CurrentDir));
         Free;
         OkCharge := TRUE;
    end;
end;

procedure TFUExport.FormClose(Sender: TObject; var Action: TCloseAction);

begin
  inherited;
    if ActionClose then exit;
    ActionClose := TRUE;
    if TPControle <> nil then
    begin
         TPControle.TOBParamFree;
         TPControle.free;
         TPControle := nil;
    end;
    TobExport.free;
    TOBVar.free;
    if Ascript <> nil then AScript.Destroy;
    GD.VidePile(FALSE);
end;

procedure TFUExport.FormCreate(Sender: TObject);
begin
  inherited;
  TPControle := TSVControle.create;
  TobExport := TOB.Create('Enregistrement', nil, -1);
  ActionClose := FALSE;
  OkCharge    := FALSE;
end;

procedure TFUExport.TableElipsisClick(Sender: TObject);
begin
  inherited;
{$IFDEF CISXPGI}
        Domaine := AglLanceFiche('CP','CPMULSCRIPT','', '', 'TRUE')  ;
        Table.text := ReadTokenPipe (Domaine, ';');
{$ELSE}
  Domaine := '';
  Table.text := MulticritereScript (nil, Domaine);
{$ENDIF}
end;


{$IFDEF CISXPGI}

procedure TFUExport.InitDelim(Codedelim : string);
var
Q                      : TQuery;
FScriptDel             : TScriptDelimite;
Stream1                : TmemoryStream;
begin
        if OkParamcount then exit;
        Stream1 := nil; FScriptDel := nil;
        Q := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+Codedelim+'" AND CIS_NATURE="X"', TRUE);
        if not Q.EOF then
	begin
            try
                Stream1 := TmemoryStream.create;
                TBlobField(Q.FindField('CIS_PARAMETRES')).SaveToStream (Stream1);
                Stream1.Seek (0,0);
                FScriptDel := LoadScriptFromStreamDelim(Stream1);
            except
                Stream1.Free;  FScriptDel.Destroy;
                Ferme(Q);
                exit;
            end;
            if OkParamcount then
            begin
                   if FListefichiers.Text = '' then FListefichiers.Text := FScriptDel.NomFichier;
            end
            else
                   if not FileExists(CurrentDonnee+'\EXPCCSX.INI') then FListefichiers.Text := FScriptDel.NomFichier;

            Stream1.Free;
	end;
     FScriptDel.free;
     Ferme(Q);
end;

procedure TFUExport.InitScript(table : string; OkC:Boolean=FALSE);
var
Q             : TQuery;
S             : TmemoryStream;
AStreamTable  : TmemoryStream;
begin
   inherited;
   S := nil;  AStreamTable := nil;
   try
              Q := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+table+'"', TRUE);
              if not Q.EOF then
              begin
                if Q.FindField('CIS_CLEPAR').asstring = 'SQL' then
                begin
                     if OkC then
                          CreateScript (TRUE, taModif, Table, FALSE);
                     ScriptSQL := TRUE;
                 end
                else
                begin
                    if Ascript <> nil then begin AScript.Destroy; Ascript:= nil end;
                    ScriptSQL := FALSE;
                    s := TmemoryStream.create;
                    TBlobField(Q.FindField('CIS_PARAMETRES')).SaveToStream (s);
                    s.Seek (0,0);
                    Ferme (Q);
                    Q := OpenSQl ('SELECT * FROM CPGZIMPCORRESP Where CIC_TABLEBLOB="CIS" and CIC_IDENTIFIANT="'+table+'"', TRUE);
                    if not Q.EOF then
                    begin
                        AStreamTable := TmemoryStream.create;
                        TBlobField(Q.FindField('CIC_OBJET')).SaveToStream (AStreamTable);
                        AStreamTable.Seek (0,0);
                    end;
                    AScript := LoadScriptFromStream(s, AStreamTable);
                    if FListefichiers.Text = '' then FListefichiers.Text := Ascript.options.Filename;
                    if Ascript.PreTrt.count <> 0 then
                    begin
                         InitDelim(table);
                    end;
                    if AStreamTable <> nil then AStreamTable.free;
                end;
              end
              else
              begin
                   PGIInfo ('Script Inexistant', 'EXPORT');
              end;
     Finally
              Ferme(Q);
              s.free;
  end;
end;

{$ELSE}

procedure TFUExport.InitDelim(Codedelim : string);
var
tbcharger              : TADOTable;
FScriptDel             : TScriptDelimite;
ABlobField             : TField;
Stream1                : TmemoryStream;
begin
  if OkParamcount then exit;
  Stream1 := nil; FScriptDel := nil;
  tbcharger := TADOTable.Create(Application);
  tbcharger.Connection := DMImport.Db as TADOConnection;
  tbcharger.TableName := DMImport.GzImpDelim.TableName;
	with tbcharger do
	begin
            if not Active then Open;
            if not Locate('CLE', Codedelim, [loCaseInsensitive]) then  exit;
            try
                ABlobField := FieldByName('PARAMETRES');
                Stream1 := TmemoryStream.create;
                TBlobField(ABlobField).SaveToStream (Stream1);
                Stream1.Seek (0,0);
                FScriptDel := LoadScriptFromStreamDelim(Stream1);
            except
                Stream1.Free;  FScriptDel.Destroy;
                Close;
                exit;
            end;
            if OkParamcount then
            begin
                   if FListefichiers.Text = '' then FListefichiers.Text := FScriptDel.NomFichier;
            end
            else
                   if not FileExists(CurrentDonnee+'\EXPCCSX.INI') then FListefichiers.Text := FScriptDel.NomFichier;

            Stream1.Free;
            Close;  free;
	end;

     FScriptDel.free;
end;

procedure TFUExport.InitScript(table : string; OkC:Boolean=FALSE);
var
FTable        : TADOTable;
S             : TmemoryStream;
AStreamTable  : TmemoryStream;
begin
   inherited;
   FTable := nil; S := nil;
   try
              FTable := TADOTable.Create(Application);
              FTable.Connection := DMImport.Db as TADOConnection;
              FTable.TableName := DMImport.GzImpReq.TableName;

              FTable.Open;
              if FTable.Locate('CLE', table, [loCaseInsensitive]) then
              begin
                if FTable.FieldByName('CLEPAR').asstring = 'SQL' then
                begin
                     if OkC then
                          CreateScript (TRUE, taModif, Table, FALSE);
                     ScriptSQL := TRUE;
                 end
                else
                begin
                    if Ascript <> nil then begin AScript.Destroy; Ascript:= nil end;
                    ScriptSQL := FALSE;
                    s := TmemoryStream.create;
                    TBlobField(FTable.FieldByName('PARAMETRES')).SaveToStream (s);
                    s.Seek (0,0);
                    AStreamTable := TmemoryStream.create;
                    TBlobField(FTable.FieldByName('TBLCOR')).SaveToStream (AStreamTable);
                    AStreamTable.Seek (0,0);
                    AScript := LoadScriptFromStream(s, AStreamTable);
                    if FListefichiers.Text = '' then FListefichiers.Text := Ascript.options.Filename;
                    if Ascript.PreTrt.count <> 0 then
                    begin
                         InitDelim(table);
                    end;
                    if AStreamTable <> nil then AStreamTable.free;
                end;
              end
              else
              begin
                   PGIInfo ('Script Inexistant', 'EXPORT');
              end;
     Finally
              FTable.Close;  FTable.free;
              s.free;
  end;
end;
{$ENDIF}

procedure TFUExport.BValideClick(Sender: TObject);
var
SLect,tmp,St          : string;
ii,isc,Indice,ind     : integer;
FieldName             : string;
TF1,TF2,TFV           : TOB;
okajout               : Boolean;
Rep,sortie            : string;
TableLien             : string;
SaveLect,LLi          : string;
begin
     if (not OkCharge) or (Ascript= nil)  then  TableExit(Sender);
     if (FILENAME.Text = '') and (AScript.FileDest = 1) then
              FILENAME.text := ExtractFileDir(FListefichiers.Text) + '\'+ Ascript.options.asciiFilename;

     SLect := FListefichiers.Text;
     Indice := 0;
     okajout := TRUE;
     for ii:= 1 to GD.RowCount do
     begin
         if GD.Cells[1, 1] = table.Text then okajout := FALSE;
         if GD.Cells[0, ii] <> '' then inc(Indice);
     end;
     if not okajout then exit;
     if Ascript= nil then exit;
     for ii:= 1+indice to length(SLect) do
     begin
        tmp := ReadTokenPipe (SLect, ';');
        if tmp = '' then break;
        TF := TOB.Create('fille'+IntToStr(ii), TobExport, -1);
        TF.AddChampSupValeur('Entree', tmp);
        TF.AddChampSupValeur('TableName', Table.text);
        Rep := ExtractFileDir(tmp);

        if (Filename.Text <> '') and (Ascript.ScriptSuivant = '') then  // ajout me pour avoir le nom de sortie qu'on veut
        begin
             sortie := Filename.Text;
             TF.AddChampSupValeur('Sortie', sortie);
        end
        else
        begin
              sortie := ExtractFileName(tmp);
              if copy (Sortie,1,2) = 'PG' then
                  sortie := ChangeFileExt(Copy(sortie,3,length(sortie)), '_'+IntToStr(ii-1)+'.TRA')
              else
                  sortie := ChangeFileExt(sortie, '.TRA');
              TF.AddChampSupValeur('Sortie', Rep+'\PG'+sortie);
        end;
        if Ascript.ScriptSuivant <> '' then
           SaveLect := Rep+'\PG'+sortie;

        TF.AddChampSupValeur('Effectue', '-');
        TF.AddChampSupValeur('Indice', IntTostr(ii));
        if (not OkParamcount) or (TOBVar= nil) then
        begin
              For isc:=0 To AScript.Variables.count-1 do
              begin

                         if TOBVar= nil then
                            TOBVar := TOB.Create('Enregistrement', nil, -1);
                          TFV := TOBVar.FindFirst(['Entree','TableName'], [tmp,Table.Text], FALSE);
                         if TFV = nil then
                         begin
                            TFV := TOB.Create('fille', TOBVar, -1);
                            TFV.AddChampSupValeur('Entree', tmp);
                            TFV.AddChampSupValeur('TableName', Table.text);
                         end;
                         St := AScript.Variables.Items[isc].Libelle;
                         St := ReadTokenpipe (St,'=');
                         TF1 := TFV.FindFirst(['Libelle'], [St], FALSE);
                         if TF1 = nil then
                         begin
                              TF1 := TOB.Create ('',TFV,-1);
                              TF1.AddChampSupValeur('Name',AScript.Variables.Items[isc].Name);
                              TF1.AddChampSupValeur('Libelle',St);
                              TF1.AddChampSupValeur('Text', AScript.Variables.Items[isc].Text);
                              TF1.AddChampSupValeur('Demandable', AScript.Variables.Items[isc].demandable);
                         end;
              end;
        end
        else
        begin
              For isc:=0 To AScript.Variables.count-1 do
              begin
                         St := AScript.Variables.Items[isc].Name;
                         St := ReadTokenpipe (St,'=');
                         TFV := TOBVar.FindFirst(['Entree','TableName'], [tmp,Table.Text], FALSE);
                         if TFV <> nil then
                         begin
                               TF1 := TFV.FindFirst(['Name'], [St], FALSE);
                               if TF1 <> nil then
                               TF1.putValue('Libelle', AScript.Variables.Items[isc].Libelle);
                         end;
              end;
        end;

        For isc := 0 to AScript.Champ.count-1 do
        begin
               FieldName := AScript.Champ[isc].Name;
               Ind := pos('_', FieldName);
               if (Ind <> 0) then
               begin
                    FieldName := copy(AScript.Champ[isc].Name, 1, Ind-1);
                    AScript.Champs[isc].NomFichExt := AScript.Name+FieldName;

                    TF1 := TF.FindFirst(['Sortief'], [AScript.Champs[isc].NomFichExt], FALSE);
                    if TF1 = nil then
                    begin
                      TF1 := TOB.Create('', TF, -1);

                      TF1.AddChampSupValeur('Entree', tmp);
                      TF1.AddChampSupValeur('TableName', Table.Text);
                      TF1.AddChampSupValeur('Sortief', AScript.Champs[isc].NomFichExt);
                      TF1.AddChampSupValeur('Effectue', '-');
                      TF1.AddChampSupValeur('FieldName', FieldName);
                      TF1.AddChampSupValeur('Lieninter', '');
                      TF1.AddChampSupValeur('SousEnsemble', FALSE);
                      TF1.AddChampSupValeur('Orderby', '');
                      TF1.AddChampSupValeur('NameScript', AScript.Name);
                      if AScript.Champ[isc].LienInter <> '' then
                         LLi := AScript.Champ[isc].LienInter;
                      if (Pos ('|', LLi) <> 0) and (Pos (FieldName+'/', LLi) <> 0) then
                      TF1.PutValue('SousEnsemble', True);


                      if AScript.Champ[isc].LienInter <> '' then
                      begin
                        TF1.PutValue('Lieninter', InterpreteVarSyntaxSt(AScript.Champ[isc].LienInter));
                        TableLien := Copy(AScript.Champ[isc].LienInter,0, Pos('/',AScript.Champ[isc].LienInter)-1);
                        TF2 := TF.FindFirst(['FieldName'], [TableLien], FALSE);
                        if TF2 <> nil then
                           TF2.PutValue('SousEnsemble', TRUE);
                       end;
                      if AScript.Champ[isc].LOrder <> '' then
                         TF1.AddChampSupValeur('Orderby', AScript.Champ[isc].LOrder);
                      //inc(Ind);
                    end;
               end
               else
               begin
                    if isc = 0 then continue;
                    if AScript.Champs[isc].NomFichExt = '' then
                         AScript.Champs[isc].NomFichExt := AScript.Name;
                    TF1 := TF.FindFirst(['Sortief'], [AScript.Champs[isc].NomFichExt], FALSE);
                    if TF1 = nil then
                    begin
                      TF1 := TOB.Create('', TF, -1);

                      TF1.AddChampSupValeur('Entree', tmp);
                      TF1.AddChampSupValeur('TableName', Table.Text);
                      TF1.AddChampSupValeur('Sortief', AScript.Champs[isc].NomFichExt);
                      TF1.AddChampSupValeur('Effectue', '-');
                      TF1.AddChampSupValeur('FieldName', FieldName);
                      TF1.AddChampSupValeur('Lieninter', '');
                      TF1.AddChampSupValeur('SousEnsemble', FALSE);
                      TF1.AddChampSupValeur('Orderby', '');
                      //inc(Ind);
                    end;
               end;
           end;
     end;
//TOBVar.SaveToFile('XXXXXX.txt',True,True,True);
    GD.ColTypes[3]:='B';
    GD.ColFormats[3]:=IntToStr(Ord(csCheckbox));
    GD.ColEditables[3]:=FALSE;

    TobExport.PutGridDetail(GD, False, False, 'Entree;Tablename;Sortie;Effectue');
    if Ascript.ScriptSuivant <> '' then
    begin
         table.Text := Ascript.ScriptSuivant;
         if TPControle <> nil then
         begin
              TPControle.TOBParamFree;
              TPControle.free;
              TPControle := nil;
         end;
         ScriptSQL := FALSE;
         TableExit(Sender);
         FListefichiers.Text := SaveLect;
         if ScriptSQL then
         begin
              inc (ii);
              TF := TOB.Create('fille'+IntToStr(ii), TobExport, -1);
              TF.AddChampSupValeur('Entree', SaveLect);
              TF.AddChampSupValeur('TableName', Table.text);
              TF.AddChampSupValeur('Effectue', '-');
              TF.AddChampSupValeur('Indice', IntTostr(ii));
              sortie := Filename.Text;
              TF.AddChampSupValeur('Sortie', sortie);
              TobExport.PutGridDetail(GD, False, False, 'Entree;Tablename;Sortie;Effectue');
         end
         else
         BValideClick(Sender);
    end;
end;


procedure TFUExport.BValiderClick(Sender: TObject);
var
iet,N,ipos,id : integer;
Name,Namesave : string;
FichierIE     : TextFile;
FTable        : TADOTable;
FFScript      : TScript;
TF,TF1,TFV,TF2: TOB;
Fichier       : string;
Opens         : Boolean;
OkResult      : Boolean;
FDefvariable  : TFDefvariable;
FLec          : TextFile;
Chaine,OrderBy: string;
OkRep,Io,isc  : integer;
NumDos,NumExe : string;
FieldName     : string;
TableLien,Lien: string;
ListeFam      : HTStringList;
FicIni        : TIniFile;
ListeLien     : TOB;
St, St1, St2  : string;
NN,LLi        : string;
TOBExp        : TOB;
TEcr,TReleve  : TOB;
TSV           : TSVControle;
{$IFDEF CISXPGI}
CpteGeneral   : string;
RefPointage   : string;
NumReleve     : integer;
{$ENDIF}
  procedure CreateListeFam;
  var
{$IFNDEF CISXPGI}
  Q1    : TQuery;
{$ENDIF}
  i     : integer;
  begin
     if Profile.Text='' then exit;
     ListeFam := HTStringList.create;
     if OkParamcount and (GetInfoVHCX.Mode <> 'I') and (FileExists(Profile.Text)) then   // si lancement par la ligne de commande
     begin
              FicIni        := TIniFile.Create(Profile.Text);
              i :=1;
              repeat
                     Name := FicIni.ReadString ('CORRESP', 'CORRESP'+IntTostr(i), '');
                     ListeFam.add(Name);
                     inc(i);
               until Name='';
               FicIni.free;
     end
     else
     begin
{$IFNDEF CISXPGI}
              Q1 := OpenSQLADO ('SELECT ChampCode from '+DMImport.GzImpCorresp.TableName+ ' Where Profile="'+ Profile.text+'"', DMImport.Db);
              while  not Q1.EOF  do
              begin
                  ListeFam.add(Q1.FindField ('ChampCode').asstring);
                  Q1.next;
              end;
              Q1.close;
{$ENDIF}
     end;
  end;
  procedure CorrespEchange;
  var
  ii, N    : integer;
  Tmp, Fam : string;
  begin
                      if OkParamcount and (GetInfoVHCX.Mode <> 'I') and (FileExists(Profile.Text)) then   // si lancement par la ligne de commande
                      begin
                        if ListeFam <> nil then
                            For ii:= 0 to  ListeFam.count-1 do
                            begin
                                 Tmp := ListeFam.Strings[ii];
                                 Fam := ReadTokenpipe (Tmp,'=');
                                 if Tmp = '' then continue;
                                 For N := 0 to  AScript.Champ.count-1 do
                                 begin
                                       if (UpperCase(AScript.Champ[N].FFamilleCorr) = Uppercase(Fam)) then
                                       begin
                                                if (Tmp <> '') and FileExists(Tmp) then
                                                begin
                                                     AScript.Champ[N].TableExterne := TRUE;
                                                     AScript.Champ[N].NomTableExt := Tmp;
                                                     if AScript.Champ[N].TableCorr = nil then
                                                     AScript.Champ[N].TableCorr := AScript.NewTableCorr;
	                                             AScript.Champ[N].TableCorr.FAssociee := true;
                                                     AScript.Champ[N].TableExist := True;
                                                end;
                                       end;  //pos
                                 end;
                            end; //for
                            Profile.Text := '';   ListeFam.free; ListeFam := nil;
                      end; // Paramcount

  end;


begin
  inherited;
         OkResult      := TRUE;  FFscript := nil;
         DatedebutExo  := iDate1900;
         DatefinExo    := iDate1900;
         FTable        := nil;  Opens := FALSE; FicIni := nil;
         ListeLien     := nil;  TEcr := nil; TReleve := nil; TSV := nil;

         if TobExport.detail.Count = 1 then Memefichier.checked := TRUE;
         if Memefichier.checked then
         begin
                   AssignFile(FichierIE, FileName.Text) ;
                   Rewrite(FichierIE) ;
                   if IoResult<>0 then
                   begin
                              PgiBox('Impossible d''écrire dans le fichier #10'+FileName.Text,'Export') ;
                              CloseFile(FichierIE);
                              Exit ;
                   end ;
         end;
         if TOBVar <> nil then
         begin
             For N:=1 To GD.RowCount-1 do
             begin
                    TFV := TOBVar.FindFirst(['Entree','TableName'], [GD.Cells[0, N],GD.Cells[1, N]], FALSE);
                    if TFV <> nil then
                          SaisieVariable(GD.Cells[0, N], TFV);
             end;
         end;
         if not OkResult then exit;
         CreateListeFam;
         For N:=0 To TobExport.detail.Count-1 do
         begin
             Lien := '';
             TableLien := '';
             Fichier := FileName.Text;
             if (TobExport.detail.Count > 1) and (not Memefichier.checked) then
                Fichier := TobExport.detail[N].getvalue ('Sortie');
             //   Fichier := ReadTokenPipe (Fichier, '.') + '_'+IntTostr(N)+ '.TRA';
             if (TobExport.detail[N].detail.count > 0) and (not Memefichier.checked) then
             begin
                   AssignFile(FichierIE, Fichier) ;
                   Rewrite(FichierIE) ;
                   if IoResult<>0 then
                   begin
                              PgiBox('Impossible d''écrire dans le fichier #10'+FileName.Text,'Export') ;
                              CloseFile(FichierIE);
                              if Profile.Text<>'' then ListeFam.free;
                              Exit ;
                   end ;
                   //OkResult := TRUE;
             end;
             CallBackDlg := nil;
             TF := TobExport.detail[N];
             InitMove(TF.detail.Count,'') ;
             if TF.detail.count = 0 then
             begin
                  if (TobExport.detail[N].Getvalue('TableName') <> '') then
                  begin
                     ScriptSQL := FALSE;
                     Namesave := CurrentDonnee+'\'+TobExport.detail[N].getvalue('TableName')+'.txt';
                     InitScript(TobExport.detail[N].getvalue('TableName'), TRUE);
                     if ScriptSQL then
                     begin
                          Fichier := TobExport.detail[N].getvalue('SORTIE');
                          Copyfile (PChar(Namesave), Pchar(Fichier),TRUE);
                          DeleteFile (CurrentDonnee+'\'+TobExport.detail[N].getvalue('TableName')+'.txt');
                          DeleteFile (CurrentDonnee+'\'+TobExport.detail[N].getvalue('TableName')+'.sch');
                     end;
                  end;
             end;
             For iet:=0 To TF.detail.Count-1 do
             begin
                  try
                      FTable := TADOTable.Create(Application);
                      FTable.Connection := DMImport.DBDonnee as TADOConnection;
                      Name := TF.detail[iet].GetValue('Sortief');
                      FieldName := TF.detail[iet].GetValue('FieldName');

                      FTable.Tablename := Name;
                      ipos := pos('.', Name);
                      if (ipos <> 0)  then Namesave := Copy (Name, 1, ipos-1)
                      else           Namesave := Name;
                      IF (not Memefichier.checked) then
                         TF.detail[iet].PutValue('Sortief', Namesave);
                      Opens := FALSE;
{$IFNDEF DBXPRESS}
                      if (not OkParamcount) and (FTable.Exists) and (AScript.Variables.count = 0) and (Ascript.PreTrt.count = 0) then
                      begin
                            FTable.Open; Opens := TRUE;
                      end
                      else
                      begin
{$ENDIF}
                            if CallBackDlg = nil then
                            begin
                                  if Ascript <> nil then begin AScript.Destroy; Ascript:= nil end;
                                  InitScript(TF.getvalue('TableName'));
                                 if (TOBVar <> nil) and (AScript.Variables.count <> 0) then
                                 begin
                                      TFV := TOBVar.FindFirst(['Entree','TableName'], [TF.detail[iet].GetValue('Entree'),TF.detail[iet].GetValue('TableName')], FALSE);
                                        if (TFV <> nil) (*and (TFV.detail.count = AScript.Variables.count)*) then
                                        begin
                                                For id:=0 To TFV.detail.count-1 do
                                                begin
                                                     For isc:=0 To AScript.Variables.count-1 do
                                                     begin
                                                          if AScript.Variables[isc].Name = TFV.detail[id].GetValue('Name')then
                                                             AScript.Variables[isc].Text := TFV.detail[id].GetValue('Text');
                                                     end;
                                                end;

                                        end;
                                 end;
                                 // traitement délimité
                                 if Ascript.PreTrt.count <> 0 then
                                   begin
                                        try
                                           TraiteScriptDelim(TF.getvalue('TableName'),TF.getvalue('Entree'),TF.getvalue('Entree')+'FIXE' );
                                           AScript.Options.FileName   := TF.getvalue('Entree')+'FIXE';
                                        except
                                        end;
                                   end;
                            CorrespEchange;
                            end;
                            if CallBackDlg = nil then
                            begin
                                 CallBackDlg := TCallBackDlg.Create(Self);
                                 CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
                            end;

                            FFScript  := Ascript.CloneScript;

                            if FTable.active then
                               FTable.Close;

                            if (pos('_', AScript.Champ[AScript.Champ.count-1].Name) <> 0) then
                               FFScript.Assign(AScript, TF.detail[iet].Getstring('FieldName'), Profile.text, ListeFam)
                            else
                               FFScript.Assign(AScript);

                            FFScript.ExecuteModeTest := TRUE;

                            FFScript.FEcrAna := AScript.EcrVent;
                            FFScript.FEcrEch := AScript.EcrEch;

                            if FFScript.FileDest = 0 then
                               FFScript.ParName := Name;
                            if (FFscript.PreTrt.count <> 0) and (FileExists(TF.getvalue('Entree')+'FIXE')) then
                                                      FFScript.Options.FileName   := TF.getvalue('Entree')+'FIXE'
                            else
                                                      FFScript.Options.FileName := TF.detail[iet].GetValue('Entree');
                            (* gestion des traitements comptables à brancher aussi sur ' E X E C I M P '*)
                            if FFscript.OptionsComptable.Correspondance.Count > 0 then
                                for Io := 0 to FFscript.OptionsComptable.correspondance.count -1 do
                                    with FFscript.OptionsComptable do
                                        DecoupeCritere(TInfoCpt(correspondance.Items[Io]).Aux, TInfoCpt(correspondance.Items[Io]).listFourchette);

                            (* Verification de la syntaxe *)

                            FFScript.ConstruitVariable(FALSE);
                            FFScript.InterpreteSyntax;
                            FFScript.ASCIIMODE := 0;

                            FFScript.Variable.Clear;
                            AjouteVariableALias(FFScript.Variable);
                            if (FFScript.Champ[0].LOrder <> '') and (AScript.FileDest = 0) then
                               OrderBy := FFScript.Champ[0].LOrder
                            else
                               OrderBy := '';
                            if OkParamcount or (FFScript.FileDest = 0) or ((AScript.FileDest = 1)  and ( not FileExists(CurrentDonnee+'\'+FFScript.ParName+'.txt')))then
                            begin
                              try FFScript.Executer(CallBackDlg.LaCallback);
                              except
                                          PGIBox('L''importation ne s''est pas correctement déroulée ','');
                                          CallBackDlg.Free;
                                          if FFscript <> nil then FFScript.destroy;
                                          FTable.Free;
                                          CloseFile(FichierIE);
                                          FiniMove ;
                                          if Profile.Text<>'' then ListeFam.free;
                                          Exit;
                              end;
                              if (AScript.FileDest = 0) then
                              begin
                               Try
                                FTable.Open;
                                EXCEPT
                                On E: Exception do
                                  ShowMessage ( 'Error : ' + E.Message ) ;
                                end ;
                              end;
                            end;
                            Namesave := FFScript.ParName;
                            if (AScript.FileDest = 1) and (iet < TF.detail.Count) then
                               TF.detail[iet].PutValue('Sortief', Namesave);

{$IFNDEF DBXPRESS}
                    end;
{$ENDIF}
                            if Opens then MoveCur(FALSE) ;
                            if (AScript.FileDest = 1)  then
                            ExportFichierText (FichierIE, CurrentDonnee+'\'+Namesave+'.txt');
                            if (FFscript <> nil) and (not Opens) then FFScript.destroy;
                            FTable.Free;
                            GD.Cells[5, iet] := FormatDateTime('hh:nn:dd/mm/yyyy',NowH);
                 Except
                             PGIBox('L''exportation ne s''est pas correctement déroulée','');
                             if (FFscript <> nil) and (not Opens) then FFScript.destroy;
                             FTable.Free;
                 end;

             end;
             if (TobExport.detail[N].detail.count > 0) and (not Memefichier.checked) then
             begin
                  CloseFile(FichierIE);
                  if OkParamcount then
                           FileClickzip (Fichier)
                  else
                           FileClickzip (Fichier);
             end;
             FiniMove ;
             if CallBackDlg <> nil then
             begin
                  CallBackDlg.Hide;
                  CallBackDlg.Free;
             end;
             GD.Cells[3, N+1] := 'X';
         end;
{$IFDEF CISXPGI}
            if TEcr <> nil then
            begin
                    TReleve.Detail.Sort('COMPTE;CODE');
                    LanceEtatTob('E', 'ECT', 'REL', TReleve, True, False, False, nil, '', 'Intégration de relevés bancaires', False);
                    if PGiAsk ('Confirmez vous l''intégration de ce relevé ?') = mryes then
                       TEcr.InsertOrUpdateDB (TRUE);
                    TEcr.free;
                    TReleve.free;
                    //AglLanceFiche('CP','RLVVISU','','','NOMREF='+RefPointage);

                    if TSV <> nil then
                    begin
                      TSV.TOBParamFree;
                      TSV.free;
                      TSV := nil;
                    end;
            end;

{$ENDIF}
            (* Ecriture dans la base*)
            
            For N :=0 To TobExport.detail.Count-1 do
            begin
                  TOBExp := TobExport.detail[N];
                  ExecuteExportFichier (FichierIE, TOBExp, TEcr, TReleve, AScript, ModeRlV, TRUE, FALSE, TSV, DatedebutExo, DatefinExo, LPays);
            end;



         if Memefichier.checked then
         begin
              CloseFile(FichierIE);
              if OkParamcount then
                         FileClickzip(Fichier)
              else
                         FileClickzip (FileName.Text);
         end;
         if OkParamcount then
          OkRep := mrYes
         else
          OkRep := PGIAsk('Confirmez-vous la supression des fichiers intermédiaires ? ','Suppression');
         if OkRep = mrYes then
         begin
            if OkParamcount and (FileName.Text <> '') then
            begin
                 if TobExport.detail.Count <> 0 then
                 begin
                      For N:=0 To TobExport.detail.Count-1 do
                      begin
                            Name := TobExport.detail[N].GetValue('SORTIE');
                            if Name = FileName.Text  then continue;
                            if N = TobExport.detail.Count-1 then
                            begin
                               Renamefile (PChar(Name), Pchar(FileName.Text));
                            end;
                            DeleteFile (Name);
                      end;
                 end;
            end;;

(* AVOIR pour drop de la table *)
           For N:=0 To TobExport.detail.Count-1 do
              begin
                   TF := TobExport.detail[N];
                   For iet:=0 To TF.detail.Count-1 do
                   begin
                         Name := TF.detail[iet].GetValue('SORTIEF');
                         DMImport.DBDonnee.Execute('DROP TABLE ' + Name);
                         DeleteFile (TF.detail[iet].GetValue('ENTREE')+'FIXE');
                   end;
              end;

         end;
         if Profile.Text<>'' then ListeFam.free;

         if not OkParamcount then
         begin
            // sauvegarde dernier filtre utilisé
            FicIni        := TIniFile.Create(CurrentDonnee+'\EXPCCSX.INI');
            FicIni.WriteString ('EXPORT', 'NOMFICHIER', FileName.Text);
            if Memefichier.checked then FicIni.WriteString ('EXPORT', 'MONOTRAITEMENT', '1')
            else FicIni.WriteString ('EXPORT', 'MONOTRAITEMENT', '0');

            if Bcompress.checked then FicIni.WriteString ('EXPORT', 'COMPRESS', '1')
            else FicIni.WriteString ('EXPORT', 'COMPRESS', '0');
            FicIni.WriteString ('EXPORT', 'SCRIPT', Table.Text);
            FicIni.WriteString ('EXPORT', 'LISTEFICHIER', FListefichiers.Text);
            FicIni.WriteString ('EXPORT', 'PROFILE', PROFILE.TEXT);

            FicIni.free;
            PgiInfo('Traitement terminé','Export') ;
         end;
end;

procedure TFUExport.GDDblClick(Sender: TObject);
var
N,iet     : integer;
FieldName : string;
TF        : TOB;
ListeEnreg: HTStringList;
//UnEnreg   : Boolean;
begin
  inherited;
  exit;
              if TobExport.detail.Count = 1 then Memefichier.checked := TRUE;
              ListeEnreg := HTStringList.Create;
              For N:=0 To TobExport.detail.Count-1 do
              begin
                  TF := TobExport.detail[N];
                  if Tf.getvalue ('Sortie') = GD.Cells[2, SavRow] then
                  begin
                     For iet:=0 To TF.detail.Count-1 do
                     begin
                           FieldName := TF.detail[iet].getvalue ('FieldName');
                           if not Memefichier.checked then
                              FieldName := FieldName+ Tf.getvalue ('Indice');
                           if ListeEnreg.IndexOf (FieldName) < 0 then
                              ListeEnreg.add(FieldName);
                     end;
                  end;
              end;
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}
              if UnEnreg then
                  ScriptVisu (GD.Cells[1, SavRow], nil, FALSE,0)
              else
                  ScriptVisu (GD.Cells[1, SavRow], ListeEnreg, TRUE,0);
{$ENDIF}
{$ENDIF}
              ListeEnreg.free;
end;

procedure TFUExport.GDSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  inherited;
  SavCol := ACol;
  SavRow := ARow;
end;

procedure TFUExport.bDefaireClick(Sender: TObject);
begin
  inherited;
  GD.VidePile(FALSE);
  TobExport.free;
  TobExport := TOB.Create('Enregistrement', nil, -1);
end;

procedure TFUExport.FileClickzip (Filen : string);
var
  FileNamec      : String ;
  Commentaire    : String ;
  TheToz         : TOZ;
  Password       : string;
  Filearchive    : string;
begin
  Password := '';
  if not BCompress.checked then exit;
    // Récupération du nom du fichier a insérer
    //
  FileNamec := FILENAME.Text;
  Filearchive := ReadTokenPipe (Filenamec, '.');
  Filearchive := Filearchive + '.zip';

  TheToz := TOZ.Create ;
  try
    if TheToz.OpenZipFile ( Filearchive, moCreate ) then
    begin
        if TheToz.OpenSession ( osAdd ) then
        begin
            if TheToz.ProcessFile (Filen, Commentaire ) then
              begin
              TheToz.CloseSession ;
              end
              else
              begin
              HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileNamec ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
              TheToz.CancelSession ;
              end ;
        end
        else
        HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileNamec ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
    end
    else
    begin
      HShowMessage ( '0;Erreur;Erreur création du fichier archive : ' + 'archive.zip' + ' impossible;E;O;O;O', '', '' ) ;
      Exit ;
    end ;
 EXCEPT
    On E: Exception do
      begin
      ShowMessage ( 'TozError : ' + E.Message ) ;
      TheToz.Free;
      end ;
 END ;
end;


function TFUExport.SaisieVariable (fich : string; TV : TOB): Boolean;
var
Ed1           : THCritMaskEdit;
Lb            : THLabel;
St            : string;
Lasth,Lastlh  : integer;
id            : integer;
Demandable    : Boolean;
begin
  Result := TRUE;
  FSaisieVar := TFVierge.Create(nil);
  FSaisieVar.BValider.Onclick := BValideClickVar;
  with FSaisieVar do
      begin
      Caption := 'Saisie des variables :' + fich;
      BorderStyle := bsDialog;
      Position := poScreenCenter;
      ClientWidth := 324;
      ClientHeight := 200+(TV.detail.count*3);
      Lasth := 21;
      Lastlh := 25;
      left := length(TV.detail[0].getvalue('Libelle'));
      For id:=0 To TV.detail.count-1 do
      begin
       if left > length(TV.detail[id].getvalue('Libelle')) then
          left := length(TV.detail[id].getvalue('Libelle'));
      end;
      Demandable := FALSE;
      For id:=0 To TV.detail.count-1 do
      begin
          if (not TV.detail[id].getvalue('demandable')) then continue;
          Demandable := TRUE;
          Lb :=  THLabel.create(nil);
          Lb.Height := 13;
          Lb.Left := 12;
          Lb.Top := Lastlh;
          Lb.Width := length(TV.detail[id].getvalue('Libelle'));
          St := TV.detail[id].getvalue('Name');
          St := ReadTokenpipe (St,'=');

          Lb.caption := TV.detail[id].getvalue('Libelle');
          Lb.Parent := FSaisieVar;

          Ed1 :=  THCritMaskEdit.create(nil);
          Ed1.Name := St;
          Ed1.tag := 100;
          Ed1.Height := 21;
          Ed1.Left := 130+ left;
          Ed1.Top := Lasth;
          Ed1.text := TV.detail[id].getvalue('Text');
          Ed1.Width := length(TV.detail[id].getvalue('Text'))+Ed1.Left;
(*          case (AScript.Variables.Items[id].TypeVar) of
          0 :; // numérique
          1 :; // alpha
          end;
*)
          Ed1.Parent := FSaisieVar;
          Lasth:=Ed1.Top+Ed1.Height+2;
          Lastlh:=Ed1.Top+Ed1.Height+5;
      end;
      try
          if (Demandable) and (ShowModal <> mrOk) then
               Result      := FALSE;
      finally
          Free;
      end;
      end;

end;


procedure TFUExport.BValideClickVar(Sender: TObject);
var
id,idv    : integer;
TL        : TControl ;
TF        : TOB;
NN        : string;
begin
  inherited;
      NN := FSaisieVar.caption;
      ReadTokenpipe (NN,':');
      TF := TOBVar.FindFirst(['Entree'], [NN], FALSE);
      if (TF = nil) or (TF.detail.count = 0) then exit;

      For id:=0 To TF.detail.count-1 do
      begin
          for idv:=0 to FSaisieVar.ControlCount-1 do
          BEGIN
               TL:=FSaisieVar.Controls[idv] ;
               If (TL is THCritMaskEdit) and (TL.Name = TF.detail[id].GetValue('Name')) and (TL.tag = 100) Then
               BEGIN
               TF.detail[id].PutValue('Text',THCritMaskEdit(TL).Text);
               break;
               END ;
          END;
      end;
end;

{$IFNDEF CISXPGI}
procedure TFUExport.savetofile(AScriptName : string);
var
	Stream1, AStreamTable : TmemoryStream;
	ABlobField : TField;
	ATblCorField : TBlobField;
	Charger : TADOTable;
begin

    if not Assigned(DMImport) then DMImport := TDMImport.Create(Application);
    Charger := DMImport.GzImpReq;
    {$IFDEF  DBXPRESS}
          DMImport.Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ AScriptName +'"';
          DMImport.Cmd.Execute;
    {$ENDIF}

    with Charger do begin
            if not Active then Open;
            {$IFDEF  DBXPRESS}
            if not EOF then begin
            {$ELSE}
            if not FindKey([AScriptName]) then begin
            {$ENDIF}
                showMessage('La requête '+aScriptName+' n''est pas une requête personnalisée');
                exit;
            end;
            ABlobField := FieldByName('PARAMETRES');
            ATblCorField := FieldByName('TBLCOR') as TBlobField;
            Stream1 := TmemoryStream.create;
            TBlobField(ABlobField).SaveToStream (Stream1);
            Stream1.Seek (0,0);

            AStreamTable := TmemoryStream.create;
            TBlobField(ATblCorField).SaveToFile ('TOTO');
            AStreamTable.Seek (0,0);
    end;
    Charger.free;
end;
{$ENDIF}

procedure TFUExport.TableExit(Sender: TObject);
begin
  inherited;
  if ActionClose then exit;
  if Table.text = '' then exit;
  if TPControle <> nil then
       TPControle.ChargeTobParam(Domaine, FALSE);
  if Ascript <> nil then begin AScript.Destroy; Ascript:= nil end;
  InitScript(Table.text);
  if (AScript <> nil) and (AScript.Shellexec <> '') then
     ExecuteShell(AScript.Shellexec);
  OkCharge    := TRUE;
end;

procedure TFUExport.FormShow(Sender: TObject);
var
Listf, tmp             : string;
sTempo                 : string;
FicIni                 : TIniFile;
i                      : integer;
Name,St                : string;
TFV,TF1                : TOB;
QP                     : Tquery;
begin
  inherited;
     if OkParamcount and (GetInfoVHCX.Mode <> 'I') then  // si lancement par la ligne de commande
     begin
          if Commande = '' then
          begin
              Commande := ParamStr(1);
              tmp := ReadTokenPipe (Commande, ';');
              CurrentDir := ReadTokenPipe (Commande, ';');
              if CurrentDir = '' then CurrentDir := ParamStr(1);
          end
          else
          CurrentDir := Commande;
          if (Copy(currentDir, 1, 5) = '/INI=') then
          begin
               sTempo := Copy(currentDir, 6, length(CurrentDir));
               if not FileExists(sTempo) then
               begin
                    PGIInfo ('Le fichier '+ sTempo+ ' n''existe pas','');
                    exit;
               end;
               FicIni        := TIniFile.Create(sTempo);

               if FicIni.SectionExists ('CORRESP') then
                  Profile.text    := sTempo;

               currentDir    := GetInfoVHCX.Directory;
               ChDir(currentDir);
               FileName.Text := CurrentDir + '\'+ GetInfoVHCX.NomFichier;
               Memefichier.checked := (GetInfoVHCX.Monotraitement='1');
               Bcompress.checked := (GetInfoVHCX.compress='1');
               Table.Text := GetInfoVHCX.Script; //FicIni.ReadString ('COMMANDE', 'SCRIPT', '');
               if Table.text = '' then
               begin
{$IFDEF CISXPGI}
                    tmp := AglLanceFiche('CP','CPMULSCRIPT','', '', 'TRUE')  ;
                    Table.text := ReadTokenPipe (tmp, ';');
{$ELSE}
                    Domaine :=GetInfoVHCX.Domaine;
                    Table.text := MulticritereScript (nil, Domaine);
{$ENDIF}
               end;
               if (Domaine = '') then Domaine := GetInfoVHCX.Domaine;
               if (Domaine = '0')  then
               begin
{$IFDEF CISXPGI}
                  QP := OpenSQL ('SELECT CIS_CLE, CIS_DOMAINE FROM CPGZIMPREQ WHERE CIS_CLE="'
                  +Table.Text+'"', TRUE);
                  if not QP.EOF then
                     Domaine := QP.FindField('CIS_Domaine').asstring;
                  Ferme (QP);
{$ELSE}
                  QP := OpenSQLADO ('Select CLE, DOMAINE from '+DMImport.GzImpReq.TableName + ' Where CLE="'+Table.Text+'"', DMImport.Db);
                  if not QP.EOF then
                               Domaine := QP.FieldByName('Domaine').asstring;
                  QP.close;
{$ENDIF}
               end;

               if Table.Text = '' then
               begin
                    FicIni.free; exit;
               end;
               Listf := GetInfoVHCX.ListeFichier;
               if Listf = '' then
               begin
                     FListefichiersElipsisClick(Sender);
                     Listf := FListefichiers.Text;
               end
               else
               begin
                   tmp := Listf;
                   sTempo := ExtractFileDir (tmp);
                   if sTempo = '' then
                      FListefichiers.Text := CurrentDir + '\'+ Listf
                   else
                      FListefichiers.Text := Listf;
               end;
               i :=1;
               repeat
                     tmp := FicIni.ReadString ('VARIABLE', 'VARIABLE'+IntTostr(i), '');
                     inc(i);
                     if tmp <> '' then
                     begin
                          if TOBVar= nil then TOBVar := TOB.Create('Enregistrement', nil, -1);
                          TFV := TOBVar.FindFirst(['Entree','TableName'], [FListefichiers.Text,Table.Text], FALSE);
                          if TFV = nil then
                          begin
                               TFV := TOB.Create('fille', TOBVar, -1);
                               TFV.AddChampSupValeur('Entree', FListefichiers.Text);
                               TFV.AddChampSupValeur('TableName', Table.text);
                          end;
                          St := tmp;
                          Name := ReadTokenpipe (St,'=');
                          TF1 := TFV.FindFirst(['Name'], [St], FALSE);
                          if TF1 = nil then
                          begin
                                TF1 := TOB.Create ('',TFV,-1);
                                TF1.AddChampSupValeur('Name',Name);
                                TF1.AddChampSupValeur('Libelle','');
                                TF1.AddChampSupValeur('Text', St);
                                TF1.AddChampSupValeur('Demandable', FALSE);
                          end;
                     end;
               until tmp='';
               FicIni.free;
          end
          else
          begin
                tmp := ReadTokenPipe (Commande, ';');
                FileName.Text := CurrentDir + '\'+tmp;
                Memefichier.checked := (ReadTokenPipe (Commande, ';')='1');
                Bcompress.checked := (ReadTokenPipe (Commande, ';')='1');
                Table.Text := ReadTokenPipe (Commande, ';');
                tmp := ReadTokenPipe (Commande, ';');
                sTempo := ExtractFileDir (tmp);
                if sTempo = '' then
                  FListefichiers.Text := CurrentDir + '\'+ tmp
                else
                   FListefichiers.Text := tmp;
          end;
          if Table.Text = '' then exit;
          if Listf = '' then exit;
          OkCharge    := FALSE;
          BValideClick (Sender);
          if  (FListefichiers.Text <> '') and (not ScriptSQL) then
              BValiderClick (Sender);
     end
     else FicheSortie := '';
     if  OkParamcount and (GetInfoVHCX.Mode = 'I') then Domaine := GetInfoVHCX.Domaine;
     HMTrad.ResizeGridColumns(GD);
     if (not OkParamcount) and (FileExists(CurrentDonnee+'\EXPCCSX.INI')) then
     begin
            // restauration dernier filtre utilisé
            FicIni        := TIniFile.Create(CurrentDonnee+'\EXPCCSX.INI');
            FileName.Text := FicIni.ReadString ('EXPORT', 'NOMFICHIER', '');
            Memefichier.checked := (FicIni.ReadString ('EXPORT', 'MONOTRAITEMENT', '')='1');
            Bcompress.checked := (FicIni.ReadString ('EXPORT', 'COMPRESS', '')='1');
            Table.Text := FicIni.ReadString ('EXPORT', 'SCRIPT', '');
            FListefichiers.Text :=  FicIni.ReadString ('EXPORT', 'LISTEFICHIER', '');
            Profile.Text        :=  FicIni.ReadString ('EXPORT', 'PROFILE', '');
            FicIni.free;
            if (Domaine = '')  then
            begin
{$IFDEF CISXPGI}
                  QP := OpenSQL ('SELECT CIS_CLE, CIS_DOMAINE FROM CPGZIMPREQ WHERE CIS_CLE="'
                  +Table.Text+'"', TRUE);
                  if not QP.EOF then
                     Domaine := QP.FindField('CIS_Domaine').asstring;
                  Ferme (QP);
{$ELSE}
                 Domaine := RenseigneDomaine (Table.Text);
                 VHCX^.Domaine := Domaine;
{$ENDIF}
            end;
     end;

end;

procedure TFUExport.TableClick(Sender: TObject);
begin
  inherited;
OkCharge    := FALSE;
end;

procedure TFUExport.ProfileElipsisClick(Sender: TObject);
begin
  inherited;
{$IFNDEF CISXPGI}
Profile.Text := SelectCorresp( 'Liste de correspondance');
{$ENDIF}
end;

procedure ExecuteExportFichier (var F : TextFile; var TOBExp, TEcr, TReleve : Tob; FScript : TScript; ModeRlV : string; Rlv, Etebac : Boolean; TSV : TSVControle;  var DatedebutExo, DatefinExo : TDateTime; Lepays : string);
var
iet                          : integer;
LLi, TableLien               : string;
ST, St1, ST2, NN             : string;
TF, TF1                      : TOB;
ListeLien                    : TOB;
CpteGeneral                  : string;
{$IFDEF CISXPGI}
RefPointage                  : string;
NumReleve                    : integer;
{$ENDIF}
begin
            ListeLien := nil;   CpteGeneral  := '';
            InitMoveProgressForm (nil, 'Export Fichier', 'Traitement en cours ...', TOBExp.detail.Count, TRUE, TRUE) ;

{$IFNDEF CISXPGI}
            if GetInfoVHCX.Domaine = '' then
               VHCX^.Domaine := RenseigneDomaine (FScript.Name);
{$ENDIF}
            For iet:=0 To TOBExp.detail.Count-1 do
            begin
                          Application.ProcessMessages;

                          if TOBExp.detail[iet].GetValue('SousEnsemble') then continue;

                          if (FScript.FileDest = 0) then
                             NN := TOBExp.detail[iet].getvalue('Sortief')
                          else
                             NN := TOBExp.detail[iet].getvalue('Sortief');
                          MoveCurProgressForm (NN);

                          if (TOBExp.detail[iet].GetValue('Lieninter') <> '') and (FScript.FileDest = 0) then
                          begin
                               LLi := TOBExp.detail[iet].GetValue('Lieninter');
                              (* Liste de lien est sous form
                                Table/Champ1,Champ2|Table2/Champ1,Champ2
                                Le TOB est structuré
                                Lien -> fille1 qui contient LISTELIEN=Table/Champ1
                                                            TABLELIEN=Script+Table+'.DB'
                                                            XWhere = Where Champ="xxx"
                                                            TableName= Table_Champ1  etc...
                                *)
                               if Pos ('|', LLi) <> 0 then
                               begin
                                     St := ReadTokenPipe(LLi,'|');
                                     ListeLien := TOB.Create ('', nil, -1);
                                     While (St <> '') do
                                     begin
                                          TF := TOB.Create('fille', ListeLien, -1);
                                          TF.AddChampSupValeur('ListeLien', St);
                                          TableLien :=TOBExp.detail[iet].GetValue('NameScript')+Copy(St, 0, Pos('/',St)-1);
                                          TF.AddChampSupValeur('TableLien', TableLien);
                                          TF.AddChampSupValeur('XWhere', '');
                                          TF.AddChampSupValeur('TableName', '');
                                          St1 := Copy(St,Pos('/',St)+1,length(St)-2);
                                          St2 := ReadTokenPipe(St1,',');
                                          While (St2 <> '') do
                                          begin
                                              TF1 := TOB.Create('fille', TF, -1);
                                              TF1.AddChampSupValeur(St2, St2);
                                              St2 := ReadTokenPipe(St1,',');
                                          end;
                                          St := ReadTokenPipe(LLi,'|');
                                     end;
                               end
                               else
                                     TableLien :=TOBExp.detail[iet].GetValue('NameScript')+Copy(TOBExp.detail[iet].GetValue('Lieninter'), 0, Pos('/',TOBExp.detail[iet].GetValue('Lieninter'))-1)

                          end
                          else
                              TableLien := '';
{$IFDEF CISXPGI}
                          if ModeRlV ='IMPORT' then
                             EcritureDanslaBase(NN, TOBExp.detail[iet].getvalue('FieldName'), GetInfoVHCX.Domaine, TEcr, TReleve, CpteGeneral, RefPointage, NumReleve,  TSV, '', '', Rlv, Etebac)
                          else
{$ENDIF}
                          begin
                              EcritureFichierExport(F, NN, TOBExp.detail[iet].getvalue('FieldName'), GetInfoVHCX.Domaine, DatedebutExo, DatefinExo, TOBExp.detail[iet].GetValue('Lieninter'), '', TableLien, TOBExp.detail[iet].getvalue('Orderby'), Lepays, ListeLien);
                              ListeLien.free; ListeLien := nil;
                          end;
            end;
            FiniMoveProgressForm;
end;

end.
