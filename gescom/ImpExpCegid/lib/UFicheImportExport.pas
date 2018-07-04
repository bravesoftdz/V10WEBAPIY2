unit UFicheImportExport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, StdCtrls, Mask, Hctrls, ComCtrls, TntComCtrls,
  UTOF,
  UTOZ,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1,
  HMsgBox,
  ed_tools,
  DateUtils,
  UTOB;

type
  TFexpImpCegid = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BImprimer: TToolbarButton97;
    PGCTRL: THPageControl2;
    TPSHTCAR: TTabSheet;
    TBSHTCONTROL: TTabSheet;
    Trace: TListBox;
    EXPORTFIC: THCritMaskEdit;
    Label1: TLabel;
    NOMFIC: THCritMaskEdit;
    Label2: TLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    RDTEXPORT: TRadioButton;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label4: TLabel;
    RDTIMPORT: TRadioButton;
    GroupBox1: TGroupBox;
    Label8: TLabel;
    IMPORTFIC: THCritMaskEdit;
    CHKZIP: TCheckBox;
    GroupBox2: TGroupBox;
    Label9: TLabel;
    CHKREPERT: TCheckBox;
    IMPORTDIR: THCritMaskEdit;
    CHKMSGSTRUCT: TCheckBox;
    CHKSTOPONERROR: TCheckBox;
    CBVIDAGEEXP: TCheckBox;
    CBVIDAGEIMP: TCheckBox;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RDTEXPORTClick(Sender: TObject);
    procedure CHKZIPClick(Sender: TObject);
    procedure CHKREPERTClick(Sender: TObject);
    procedure RDTIMPORTClick(Sender: TObject);
  private
    { Déclarations privées }
    ExpDir,ExpFile : string;
    procedure ExportDossier;
    procedure ImportDOS(FileN,ExportN: string);
    function RendSQLTable : string;
    function ConstitueRequete(maTable: String): string;
    function PGCreatZipFile(Archive: string; CodeSession: TModeOpenZip;var TheTOZ: TOZ; Ecran: TFORM): Boolean;
    function TraitementTable(TheTOZ: TOZ; MaTable, Prefixe: string;Numversion: integer; DDebut, DFin: TdateTime;CodeEx,LibEx: string; ParExercice,premier : boolean): boolean;
    function PGZipFile(Fichier: string; TheTOZ: TOZ;Ecran: TFORM): Boolean;
    procedure PGFermeZipFile(TheTOZ: TOZ);
    procedure PGVideDirectory(FileN: string);
  public
    { Déclarations publiques }
  end;


procedure LanceImportExport;

implementation

uses
  CommonTools;

{$R *.dfm}

procedure LanceImportExport;
  var FexpImpCegid: TFexpImpCegid;
begin
  FexpImpCegid := TFexpImpCegid.create(Application);
  TRY
    FexpimpCegid.ShowModal;
  FINALLY
    FexpImpCegid.free;
  END;
end;

procedure TFexpImpCegid.BValiderClick(Sender: TObject);
begin
  Trace.Clear;
  if RDTEXPORT.Checked then
  begin
    ExpDir := EXPORTFIC.Text;
    ExpFile := trim(NOMFIC.text);

    if ExpDir = '' then
    begin
      PgiBox('Vous n''avez pas renseigner le réperoire de transfert !', self.caption);
      ModalResult := mrNone;
      exit;
    end;
    if ExpFile = '' then
    begin
      PgiBox('Vous n''avez pas renseigner le nom du fichier archive !', self.caption);
      ModalResult := mrNone;
      exit;
    end;

    PGCTRL.ActivePage := TBSHTCONTROL;
    ExportDossier;
  end
  else
  begin
    if PgiAsk('Attention: Cet import écrasera de manière définitive les données de ce dossier.#13#10 Voulez vous continuer le traitement ?', self.caption) <> Mryes then Exit;
    if (IMPORTFIC.text = '') and (CHKZIP.checked) then
    begin
      PgiError('Vous devez renseigner le nom du fichier à importer !', Self.caption);
      exit;
    end;
    if (IMPORTDIR.text = '') and (CHKREPERT.checked) then
    begin
      PgiError('Vous devez renseigner le nom du répertoire ou les fichiers sont stockés !', Self.caption);
      exit;
    end;
    if (PGCTRL  <> nil) and (TBSHTCONTROL <> nil) then PGCTRL.ActivePage := TBSHTCONTROL;
    IMPORTDOS(IMPORTFIC.Text,IMPORTDIR.Text);
  end;
  ModalResult := 0;
end;

procedure TFexpImpCegid.BFermeClick(Sender: TObject);
begin
  close;
end;

procedure TFexpImpCegid.BImprimerClick(Sender: TObject);
begin
//
end;

procedure TFexpImpCegid.HelpBtnClick(Sender: TObject);
begin
//
end;

procedure TFexpImpCegid.FormCreate(Sender: TObject);
begin
//
end;

procedure TFexpImpCegid.FormShow(Sender: TObject);
begin
//
	CHKZIP.Enabled := false;
  CHKREPERT.Enabled := false;
  IMPORTFIC.Enabled := false;
  IMPORTDIR.Enabled := false;
end;

procedure TFexpImpCegid.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//
end;

procedure TFexpImpCegid.RDTEXPORTClick(Sender: TObject);
begin
  CHKMSGSTRUCT.Enabled := false;
  CHKSTOPONERROR.Enabled := false;
  CHKZIP.Enabled := false;
  CHKREPERT.Enabled := false;
  CBVIDAGEIMP.Enabled := false;
  IMPORTFIC.Enabled := false;
  IMPORTDIR.Enabled := false;
  //
  CBVIDAGEEXP.Enabled := true;
  NOMFIC.Enabled := true;
  EXPORTFIC.Enabled := true;
end;

procedure TFexpImpCegid.CHKZIPClick(Sender: TObject);
begin
  if CHKZIP.Checked then
  begin
    CHKREPERT.checked := false;
    IMPORTFIC.enabled := true;
    IMPORTDIR.Enabled := false;
  end else
  begin
    CHKREPERT.checked := true;
    IMPORTFIC.enabled := false;
    IMPORTDIR.Enabled := true;
  end;

end;

procedure TFexpImpCegid.CHKREPERTClick(Sender: TObject);
begin
  if CHKREPERT.checked then
  begin
    CHKZip.Checked := false;
    IMPORTDIR.enabled := true;
  end else
  begin
    CHKZip.Checked := true;
    IMPORTDIR.enabled := false;
  end;
end;

function TFexpImpCegid.ConstitueRequete(maTable: String): string;
begin
  if maTable = 'PARAMSOC' then
  begin
    result := 'SELECT * FROM PARAMSOC WHERE '+ Tools.GetPSocTreeToExport;
  end
  else if maTable = 'CHOIXCOD' then result := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE IN (SELECT DO_TYPE FROM DECOMBOS WHERE DO_DOMAINE IN ("C","P","Y","T","D","0") AND DO_PREFIXE="CC")'
  else if maTable = 'CHOIXEXT' then result := 'SELECT * FROM CHOIXEXT WHERE YX_TYPE IN (SELECT DO_TYPE FROM DECOMBOS WHERE DO_DOMAINE IN ("C","P","Y","T","D","0") AND DO_PREFIXE="YX")'
  else result := 'SELECT * FROM ' + maTable;
end;

procedure TFexpImpCegid.ExportDossier;
var
  T_Mere, T, T1, T_EXERCICE: TOB;
  St, LibEx,Prefixe,CodeEx: string;
  Q: TQuery;
  Rep, Nbre, II,version,I : integer;
  ExportOk: Boolean;
  MaTable, ChampMvt, ZipFile,msg: string;
  TheToz: TOZ;
  Pexercice : boolean;
  DDebut,DFin : Tdatetime;
begin
  ZipFile := IncludeTrailingPathDelimiter (ExpDir) + ExpFile + '.zip';
  T_EXERCICE := TOB.Create ('LES EXERCICES',nil,-1);
  T_Mere := TOB.Create('Ma TOB', nil, -1);
  T := TOB.Create('La TOB des tables', nil, -1);

  TRY
    Q := OpenSQl ('SELECT EX_EXERCICE,EX_DATEDEBUT,EX_DATEFIN,EX_ABREGE,EX_LIBELLE FROM EXERCICE',True);
    if not Q.eof then T_EXERCICE.loadDetailDb ('EXERCICE','','',Q,false);
    ferme (Q);
    st := RendSQLTable;
    Q := OPENSQL(ST, TRUE);
    T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    //
    Trace.Items.Add('Les éléments suivants seront analysés et pris en compte :');
    for II := 0 to T.detail.count - 1 do
    begin
      T1 := T.Detail[II];
      if T1 <> nil then
      begin
        st := 'SELECT COUNT (*) NBRE FROM ' + T1.GetValue('DT_NOMTABLE');
        Q := OPENSQL(ST, TRUE);
        if not Q.EOF then
        begin
          Nbre := Q.FindField('NBRE').AsInteger;
          if Nbre > 0 then Trace.Items.Add('Table ' + T1.GetValue('DT_LIBELLE'));
        end;
        Ferme(Q);
      end;
    end;
    TheTOZ := nil;
    //
    if not PGCreatZipFile(ZipFile, moCreate, TheTOZ, self) then
    begin
      T.free;
      T_Mere.free;
      Trace.Items.Add('Erreur sur fichier archive : ' +ZipFile);
      exit;
    end;

    Rep := PgiAsk('Voulez vous continuer le traitement', self.Caption);
    if Rep <> MrYes then
    begin
      exit;
    end;
    InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', T.detail.count, FALSE, TRUE);

    ExportOk := TRUE;
    for I := 0 to T.detail.count - 1 do
    begin
      T1 := T.Detail[I];
      if T1 <> nil then
      begin
        MaTable := T1.GetValue('DT_NOMTABLE');
        Prefixe :=  T1.GetValue('DT_PREFIXE');
        version := T1.GetInteger('DT_NUMVERSION');
        //
        St := ConstitueRequete (maTable);
        Pexercice := FALSE;
        if (MaTable <> 'EXERCICE') and (MaTable <> 'MENU') then
        begin
          ChampMvt := 'SELECT 1 FROM DECHAMPS WHERE DH_NOMCHAMP = "' + Prefixe + '_DATECREATION"';
          if (ExisteSQL(ChampMvt)) then Pexercice := TRUE;
        end;
        if Pexercice then
        begin
          for II := 0 to  T_EXERCICE.detail.count -1 do
          begin
            DDebut := T_EXERCICE.detail[II].GetDateTime('EX_DATEDEBUT');
            if II = T_EXERCICE.detail.count -1 then DFin := iDate2099
                                               else DFin := T_EXERCICE.detail[II].GetDateTime('EX_DATEFIN');
            LibEx :=  T_EXERCICE.detail[II].GetString('EX_LIBELLE');
            CodeEx := T_EXERCICE.detail[II].GetString('EX_EXERCICE');
            //
            if not TraitementTable(TheTOZ,MaTable,Prefixe,version,DDebut,DFin,CodeEx,LibEx,PExercice,(II=0)) then
            begin
              Msg := 'Erreur dans table '+MaTable+' Exercice ' + LibEx;
              PgiError(Msg, self.caption);
              Trace.Items.add(Msg);
            end;
          end;
        end else
        begin
          DDebut := IDate1900;
          DFin := iDate2099;
          if not TraitementTable(TheTOZ,MaTable,Prefixe,version,DDebut,DFin,'',LibEx,false,true) then
          begin
            Msg := 'Erreur dans table '+MaTable;
            PgiError(Msg, self.caption);
            Trace.Items.add(Msg);
          end;
        end;
      end;
    end;
    MoveCurProgressForm('Fin de traitement des données');
    if not Exportok then PgiError('Le traitement est abandonné', self.caption)
                    else PGFermeZipFile(TheTOZ);
    if CBVIDAGEEXP.checked then PGVideDirectory(ExpDir);
    if Assigned(TheToz) then TheToz.Free;
    try
      Trace.Items.add('Ecriture du fichier OK');
    except
      PGIBox('Une erreur est survenue lors de l''écriture du fichier', self.caption);
      Trace.Items.add('Une erreur est survenue lors de l''écriture du fichier');
    end;
    FiniMoveProgressForm();
    PgiBox('Exportation terminée', self.caption);
  FINALLY
    T.free;
    T_Mere.free;
    T_EXERCICE.free;
  END;
end;

procedure TFexpImpCegid.ImportDOS(FileN,ExportN: string);
var
  rep,  Numvers, ret, LaVersion: Integer;
  T_Mere, T, T1, T2 : TOB;
  Q: TQUERY;
  Repert, LeNom,  st, MonFic: string;
  TheTOZ: TOZ;
  sr: TsearchRec;
  OkOk : Boolean;
begin

  T_Mere := TOB.Create('Ma TOB', nil, -1);
  T := TOB.Create('La TOB des tables', nil, -1);
  TRY
    if FileN<> '' then repert := ExtractFilePath(FileN)
                  else Repert := IncludeTrailingBackslash (ExportN);
    TheTOZ := nil;
    if (FileN <> '') then
    begin
      Trace.Items.Add('Phase 1 - unzip du fichier...');

      if not PGCreatZipFile(FileN, moOpen, TheTOZ, Self) then
      begin
        PgiError('Traitement abandonné', Self.caption);
        exit;
      end
      else PGFermeZipFile(TheTOZ);
      Trace.Items.Add('Phase 1 - unzip du fichier...OK');
    end;

    ST := RendSQLTable;
    Q := OPENSQL(ST, TRUE);
    T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    InitMoveProgressForm(nil, 'Analyse des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
    ret := FindFirst(Repert + '*.SEB', 0 + faAnyFile, sr);
    Rep := MrYes;
    while (ret = 0) and (rep = mrYes) do
    begin
      //
      T_Mere.ClearDetail;
      //
      MonFic := Repert + sr.Name;
      MoveCurProgressForm('Analyse des données du fichier ' + sr.Name);
      //
      TOBLoadFromBinFile(MonFic, nil, T_Mere);
      T1 := T_Mere.Detail[0];
      if Assigned(T1) then
      begin
        LeNom := T1.GetValue('TABLE');
        LaVersion := T1.GetValue('VERSION');
        T2 := T.FindFirst(['DT_NOMTABLE'], [LeNom], FALSE);
        if T2 <> nil then
        begin
          if (T2.getString('VERIFIED')<> 'X') then
          begin
            NumVers := T2.GetValue('DT_NUMVERSION');
            if (LaVersion <> NumVers) and (NumVers <> 0) and (T2.getString('VERIFIED')<> 'X') and (CHKMSGSTRUCT.checked) then
            begin
              if PgiAsk ('ATTENTION : la structure d''origine de la table '+LeNom+' n''est pas identique à celle de ce dossier.#13#10 Confirmez-vous quand même l''import ?',self.caption) <> mryes then
              begin
                rep := MrCancel;
              end else T2.SetString('VERIFIED','X');
            end;
          end;
        end;
      end;
      ret := FindNext(sr);
    end;
    sysutils.FindClose(sr);

    if Rep <> MrYes then
    begin
      PgiError('Des anomalies de structure ont été détectées ou bien vous avez abandonné le traitement', self.caption);
      FiniMoveProgressForm();
      Trace.Items.Add('Traitement abandonné');
      exit;
    end;

    if PgiAsk ('ATTENTION : Confirmez-vous le traitement d''import des données ?',self.caption) <> mryes then
    begin
      FiniMoveProgressForm();
      Trace.Items.Add('Traitement abandonné');
      exit;
    end;
    //
    ret := FindFirst(Repert + '*.TEB', 0 + faAnyFile, sr);
    OkOk := TRUE;
    while (ret = 0) and (OkOk) do
    begin
      T_mere.ClearDetail;
      //
      MonFic := Repert + sr.Name;
      TOBLoadFromBinFile(MonFic, nil, T_Mere);
      T1 := T_Mere.Detail[0];


      MoveCurProgressForm('Traitement des données du fichier ' + sr.Name);
      try
        T2 := T.FindFirst(['DT_NOMTABLE'], [ T1.GetValue('TABLE')], FALSE);
        if T2 = nil then continue; // la table n''existe pas dans la destination....
        if (T1.getString('VIDAGE')='X') then
        begin
          if (T2.GetString('REINIT')<> 'X') then
          begin
            ExecuteSql('DELETE FROM '+T1.GetValue('TABLE'));
            T2.SetString('REINIT','X');
          end;
          T_Mere.SetAllModifie(true); 
          OkOk := T_Mere.InsertDBByNivel(TRUE);
        end else
        begin
          T_Mere.SetAllModifie(true); 
          OkOk := T_Mere.InsertOrUpdateDB(TRUE);
        end;
        Trace.Items.add('Traitement du fichier ' + sr.Name + ' OK');
      except
        Trace.Items.add('Erreur de traitement des données du fichier ' + sr.Name);
        if (CHKSTOPONERROR.checked) then
        begin
          PgiError('ATTENTION : Le traitement des données du fichier '+sr.name+' n''est pas correct.#13#10 votre base de données n''est pas utilisable en l''état !', self.Caption);
          OkOk := FALSE;
        end;
      end;
      ret := FindNext(sr);
    end;
    FiniMoveProgressForm();
    sysutils.FindClose(sr);
    if OkOk then Repert := Copy(Repert, 1, Strlen(PChar(Repert)) - 1);
    if CBVIDAGEIMP.Checked then PGVideDirectory(Repert);
    EXECUTESQL ('DELETE FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="PSA_PRENOM"');
    PgiBox('Import terminé', self.caption);
  FINALLY
    T_Mere.free;
    T.free;
  END;
end;

function TFexpImpCegid.PGCreatZipFile(Archive: string;CodeSession: TModeOpenZip; var TheTOZ: TOZ; Ecran: TFORM): Boolean;
begin
  result := false;
  TheToz := TOZ.Create;
  try
    if TheToz.OpenZipFile(Archive, CodeSession) then
    begin
      if CodeSession = moCreate then
        if TheToz.OpenSession(osAdd) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');

      if CodeSession = moOpen then
        if TheToz.OpenSession(osExt) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');
    end
    else
    begin
      HShowMessage('0;Erreur;Erreur création du fichier archive : ' + Archive + ' impossible;E;O;O;O', '', '');
      Exit;
    end;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.Caption);
      TheToz.Free;
    end;
  end;
end;

procedure TFexpImpCegid.PGFermeZipFile(TheTOZ: TOZ);
begin
  TheToz.CloseSession;
end;

procedure TFexpImpCegid.PGVideDirectory(FileN: string);
var
  sr: TsearchRec;
  ret: Integer;
  MonFic: string;
begin

  ret := FindFirst(FileN + '\*.TEB', 0 + faAnyFile, sr);
  while ret = 0 do
  begin
    MonFic := FileN + '\' + sr.Name;
    if FileExists(MonFic) then DeleteFile(PChar(MonFic));
    ret := FindNext(sr);
  end;
  sysutils.FindClose(sr);
  ret := FindFirst(FileN + '\*.SEB', 0 + faAnyFile, sr);
  while ret = 0 do
  begin
    MonFic := FileN + '\' + sr.Name;
    if FileExists(MonFic) then DeleteFile(PChar(MonFic));
    ret := FindNext(sr);
  end;
  sysutils.FindClose(sr);
end;

function TFexpImpCegid.PGZipFile(Fichier: string; TheTOZ: TOZ;Ecran: TFORM): Boolean;
begin
  result := false;
  try
    if TheToz.ProcessFile(Fichier, '') then result := TRUE;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.caption);
      TheToz.Free;
    end;
  end;
end;

function TFexpImpCegid.RendSQLTable: string;
begin
  result := 'SELECT '+
            'DT_NOMTABLE,DT_PREFIXE,DT_LIBELLE,DT_NUMVERSION,"-" AS REINIT,"-" AS VERIFIED ' +
            'FROM DETABLES '+
            'WHERE '+
            '(DT_DOMAINE IN ("C","Y","T","D","0","P")) OR '+
            '(DT_NOMTABLE = "CHOIXCOD") OR (DT_NOMTABLE = "MENU") OR '+
            '(DT_NOMTABLE = "PARAMSOC") OR '+
            '(DT_NOMTABLE = "ETABLISS") OR (DT_NOMTABLE = "ETABLCOMPL") OR (DT_NOMTABLE = "SOCIETE")';
  result := result + 'ORDER BY DT_NOMTABLE';
end;

function TFexpImpCegid.TraitementTable(TheTOZ: TOZ; MaTable,Prefixe: string; Numversion: integer; DDebut, DFin: TdateTime; CodeEx,LibEx: string; ParExercice, premier: boolean): boolean;
var SQl : String;
    QQ : TQuery;
    T_MERE,T_Fic,T_STFIC,T_STMERE : TOB;
    TheFic,TheStFic : string;
begin
  result := true;
  T_STMEre := TOB.Create ('UNE STRUCTURE',nil,-1);
  T_STFIC := TOB.Create(MaTable+'_', T_STMere, -1);
  T_STFIC.AddChampSupValeur('TABLE', MaTable);
  T_STFIC.AddChampSupValeur('VERSION', Numversion);
  T_STFIC.AddChampSupValeur('VIDAGE', 'X');
  //
  T_Mere := TOB.Create('Ma tob', nil, -1);
  T_Fic := TOB.Create(MaTable+'_', T_Mere, -1);
  T_Fic.AddChampSupValeur('TABLE', MaTable);
  T_Fic.AddChampSupValeur('VERSION', Numversion);
  T_Fic.AddChampSupValeur('VIDAGE', 'X');
  if Pos(maTable,'MENU;PARAMSOC;CHOIXCOD;CHOIXEXT;') > 0 then T_Fic.SetString('VIDAGE', '-');
  TRY
    TheStFic := ExpDir + '\' + MaTable+'.SEB';
    //
    if maTable = 'MENU' then
    begin
			SQL := 'SELECT * FROM MENU '+
      			 'WHERE '+
      			 '(mn_1=0 and mn_2 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377)) or (mn_1 in (41,42,43,44,46,47,48,49,200,303,310,347,639,371,372,373,374,375,376,377)) or '+
             '(mn_1=0 and mn_2 in (8,9,11,12,13,14,16,17,18,26,27,324,330,340,361)) or (mn_1 in (8,9,11,12,13,14,16,17,18,26,27,324,330,340,361))';
    end else
    begin
    	SQL := 'SELECT * FROM '+maTable;
    end;
    if (ParExercice) then
    begin
      MoveCurProgressForm('Traitement de la table ' + MaTable + ' ' + LibEx);
      if premier then SQL := Sql + ' WHERE '+Prefixe+'_DATECREATION >= "'+USDATETIME(IDate1900)+'" AND '+Prefixe+'_DATECREATION <= "'+USDATETIME(DFin)+'"'
                 else SQL := Sql + ' WHERE '+Prefixe+'_DATECREATION >= "'+USDATETIME(DDebut)+'" AND '+Prefixe+'_DATECREATION <= "'+USDATETIME(DFin)+'"';
      TheFic := ExpDir + '\' +MaTable + '#'+CodeEx+'.TEB';
    end else
    begin
      MoveCurProgressForm('Traitement de la table ' + MaTable );
      TheFic := ExpDir + '\' + MaTable +'.TEB';
    end;

    QQ := OPENSQL(SQL, TRUE);
    if not QQ.eof then
    begin
      T_Fic.LoadDetailDb(MaTable, '', '', QQ, FALSE);
    end;
    Ferme(QQ);
    //
    if premier then
    begin
      if FileExists(TheSTFic) then DeleteFile(PChar(TheSTFic));
      T_STFic.SaveToBinFile(TheSTFic, FALSE, TRUE, TRUE, FALSE);
      if not PGZipFile(TheSTFic, TheTOZ, self) then result := FALSE;
    end;
    //
    if FileExists(TheFic) then DeleteFile(PChar(TheFic));
    T_Fic.SaveToBinFile(TheFic, FALSE, TRUE, TRUE, FALSE);
    if not PGZipFile(TheFic, TheTOZ, self) then result := FALSE;
  FINALLY
    FreeAndNil(T_Mere);
    T_STMEre.free;
  END;
end;

procedure TFexpImpCegid.RDTIMPORTClick(Sender: TObject);
begin
  CHKMSGSTRUCT.Enabled := true;
  CHKSTOPONERROR.Enabled := true;
  CHKZIP.Enabled := true;
  CHKREPERT.Enabled := true;
  CBVIDAGEIMP.Enabled := true;
  IMPORTFIC.Enabled := true;

  //
  CBVIDAGEEXP.Enabled := false;
  NOMFIC.Enabled := false;
  EXPORTFIC.Enabled := false;
end;

end.



