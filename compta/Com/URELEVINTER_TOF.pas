unit URELEVINTER_TOF;

interface

uses
    Windows,
    Messages,
    StdCtrls,
    Controls,
    Classes,
    forms,
    ComCtrls,
    Shellapi,
    Sysutils, // StrToInt, IntToStr, Date, EncodeDate, StrToFloat, TSearchRec, FindFirst, FindNext, faAnyFile, RenameFile, ExtractFilePath, DateToStr, Uppercase
{$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    FE_Main,
    EdtREtat, {LanceEtatTob}
    db,
    Mul,
{$ELSE}
    MaineAGL,
    UtileAgl, {LanceEtatTob}
    UscriptTob,
    UtilTrans,
    eMul,
{$ENDIF}
    ULibCpContexte,
    dpTOFVisuReleve, // CPLanceFicheVisuRel
    UTOF,     // TOF
    UTOB,     // TOB
    Uscript,
    UExecute,
    UAssistConcept,
    UDMIMP,
    UControlParam,
    inifiles,
    Dialogs,  // TOpenDialog
    HCtrls,   // Contrôle Halley
    HEnt1,    // TraduireMemoire, iDate1900, iDate2099, V_PGI
    HMsgBox,  // THMsgBox
    Vierge,   // TFVierge
    Filectrl, // DirectoryExists
    HTB97,    // TToolbarButton97
    Ent1,     // VH
    UParametre,
    UExport,
    EcHg_code,
    CPMulscript,
    UlibWindows, UTOZ,
    uYFILESTD,        // BibleFichier (modes oératoires en pdf)
    ParamSoc; // GetParamSocSecur


procedure CPLanceFiche_RecupReleveInter (Mode : string= 'IMPORT');
//procedure LancementCommande;

Type
     {TOF RECUPRELEVE}
     TOF_RECUPRELEVINTER = Class (TOF)
     private
       TobReleve         : TOB;
       GD                : THGrid ;
       Repertoire        : string ;
       Extent            : string;
       ComboScript       : THValComboBox;
       SavCol,SavRow     : integer;
       ListeScripts      : TStringList;
       TSV               : TSVControle;
       ModeRlv           : string;
       FSaisieVar        : TFVierge;

       Procedure InitGrid ;
       procedure BVALIDELISTEClick(Sender: TObject);
       Procedure ScruteFichier (var Dossier: string; var filtre: string; Attributs: integer);
       procedure GDKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure GDSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
       procedure GDDblClick(Sender: TObject);
       procedure ComboScriptChange(Sender: TObject);
       function InitScriptGroupe(table : string) : TScript;
       procedure RecoFichiers(Dossier : string; Liste : TStringList; Extension : string; Attributs : integer);
       procedure TestFichierScripts( NomFichier : string; NomScript : string);
       procedure ChargementTOBExp (TOBE : TOB; var Ascript : TScript);
       procedure BReconnaissanceClick(Sender: TObject);
       procedure VisuFicDonnees (Sender: TObject);
       procedure ParamScript (Sender: TObject);
       procedure ParamScriptRec (Sender: TObject);
       procedure BValiderClick(Sender: TObject);
       procedure RepertoireOnChange(Sender: TObject);
       procedure TraiteMessage (var Msg: TMsg; var Handled: Boolean);
       procedure bDefaireClick(Sender: TObject);
       procedure Execute (TBN : string;  var TEcr,TReleve : TOB; N : integer);
       procedure BScruterClick(Sender: TObject);
       procedure bTagClick(Sender: TObject);
       procedure ChargeTob(value : string; Row : integer);
       function  SaisieVariableGroupe (fich : string; TV : TOB): Boolean;
       procedure BValideClickVarGrp(Sender: TObject);
     public
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (S : String ) ; override ;

     END ;
  const NbColonne = 5;

implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    , Constantes, Commun
    , Math, UtilPGI
    ;
procedure CPLanceFiche_RecupReleveInter (Mode : string= 'IMPORT');
begin
  if CEstPointageEnConsultationSurDossier then
  begin
    PgiInfo('Vous avez indiqué une liaison avec une comptabilité ' +
            RechDom('CPLIENCOMPTABILITE',GetParamSocSecur('SO_CPLIENGAMME',''), False) +
            ' et la gestion du pointage ' + #10 +
            'est effectuée ' + RechDom('CPPOINTAGESX', GetParamSocSecur('SO_CPPOINTAGESX',''), False) + '. ' +
            'Vous n''avez pas accès à cette commande.', 'Intégration des relevés');
    Exit;
  end;
  if Mode = 'IMPORT' then
     AGLLanceFiche('CP', 'RLVINTEGREINTER', '', '', Mode)
  else
     AGLLanceFiche('CP', 'RLVEXPORTINTER', '', '', Mode)

end;

(*procedure LancementCommande;
var
FicINI       : string;
begin
  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then gDirEchDuDos := V_PGI.DosPath + '\Echanges'
  else  gDirEchDuDos := V_PGI.DosPath + 'Echanges';
  if Not DirectoryExists (gDirEchDuDos) then
  CreateDir (gDirEchDuDos);

  FicINI := gDirEchDuDos + ConstFicIniDemande;
  SysUtils.DeleteFile (FicIni);
 CreeFicIniPourCISX ('TRA', 'XVIRINTBELGE', 'C:\devpgi\RELEVE\CISX1.txt', gDirEchDuDos+'\PGCISX1.txt', nil);
 LanceExportCisx ('/INI='+FicINI, 'EXPORT', 'BEL');
end;
*)

procedure TOF_RECUPRELEVINTER.OnArgument(S: string);
begin
     ModeRlV := S;
    TFMul(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('GFILES')));

end;

procedure TOF_RECUPRELEVINTER.OnLoad;
var
Chemindos, Cheminstd ,sTempo                          : string;
FicIni                                                : TIniFile;
QP                                                    : TQuery;
Vales                                                 : HTstringList;
CodeRetour                                            : integer;
zip                                                   : TOZ;
FichierAIntegrer                                      : string;
begin
inherited ;
//if Ecran<>nil then TFORM(Ecran).OnKeyDown:=OnKeyDown ;
TFVierge(Ecran).HelpContext:=7773000;
GD :=THGrid(GetControl('GFILES'));
TCPContexte.GetCurrent.VarCisx.CHARGECIX;
ComboScript := THValComboBox(GetControl('ComboScript'));
TToolbarButton97(GetControl('BCHERCHER')).OnClick       := BVALIDELISTEClick;
TToolbarButton97(GetControl('BRECONNAISSANCE')).OnClick := BReconnaissanceClick;
TToolbarButton97(GetControl('BVOIRRELEVE')).OnClick     := VisuFicDonnees;
TToolbarButton97(GetControl('BPARAMSCRIPT')).OnClick    := ParamScript;
TToolbarButton97(GetControl('BPARAMSCRIPTREC')).OnClick := ParamScriptRec;
TToolbarButton97(GetControl('BVALIDER')).OnClick        := BValiderClick;
TToolbarButton97(GetControl('BSCRUTER')).OnClick        := BScruterClick;
TToolbarButton97(GetControl('bDefaire')).OnClick        := bDefaireClick;
TToolbarButton97(GetControl('bTag')).OnClick            := bTagClick;
THEdit(GetControl('REPERTOIRE')).OnChange     := RepertoireOnChange;
GD.OnKeyDown                                  := GDKeyDown;
GD.OnSelectCell                               := GDSelectCell;
GD.OnDblClick                                 := GDDblClick;
ComboScript.OnChange := ComboScriptChange;
  // Liste des scripts de reconnaissance
ListeScripts := TStringList.Create;
(*if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then Chemindos := V_PGI.DosPath+'\CISX'
else  Chemindos := V_PGI.DosPath+'CISX' ;
if V_PGI.StdPath[length (V_PGI.StdPath)] <> '\' then Cheminstd := V_PGI.StdPath+'\CISX'
else  Cheminstd := V_PGI.StdPath+'CISX' ;
*)
Chemindos := GetWindowsTempPath  +'CISX';
if not DirectoryExists(Chemindos) then
  CreateDir(Chemindos);

CodeRetour :=  AGL_YFILESTD_EXTRACT (FichierAIntegrer, 'CISX', 'PGZIMPACCESSD.ZIP','C', 'COMPTA', 'PARAM',
'', '', FALSE, V_PGI.LanguePrinc, 'CEG');
if CodeRetour <> -1 then
PGIInfo ('Erreur d''accès à la base Access')
else
begin
  zip := TOZ.Create;
  zip.OpenZipFile(FichierAIntegrer, moOpen);
  if not zip.OpenSession(osExt) then
  begin
    PGIInfo('échec ouverture zip', 'OpenSession(osExtAll)');
    zip.Free;
  end
  else
  begin // extraction
           sysutils.DeleteFile(Chemindos + '\PGZIMPACCESSD.MDB');
           zip.SetDirOut(Chemindos);
           zip.CloseSession;
           zip.Free;
  end;
end;
Cheminstd := Chemindos;

InitDB (Cheminstd, Chemindos);
sTempo := CurrentDonnee+'\GEXPCCSX.INI';
if (FileExists(sTempo)) then
begin
  // restauration dernier filtre utilisé
  FicIni              := TIniFile.Create(sTempo);
  SetControlText('REPERTOIRE', FicIni.ReadString ('EXPORT', 'ENTREE', ''));
  SetControlText('EEXTENTION', FicIni.ReadString ('EXPORT', 'EXTENTION', ''));
  FicIni.free;
end;
  DragAcceptFiles (THedit(GetControl('REPERTOIRE')).handle, true);
  Application.OnMessage := TraiteMessage;
  Vales := HTStringList.Create;
  QP := OpenSQL ('SELECT CIS_CLE from CPGZIMPREQ Where CIS_NATURE<>"X" AND CIS_DOMAINE="'+GetInfoVHCX.Domaine+'"', TRUE);
  While  not (QP.Eof) do
  begin
          Vales.add(QP.FindField('CIS_CLE').asstring);
        QP.Next;
  end;
  Comboscript.Items := Vales;
  Comboscript.Values := Vales;
  Vales.clear;
  Vales.Free;
  QP.free;
  InitGrid ;
  TSV := nil;
  if ModeRlV = 'EXPORT' then
    THValComboBox(GetControl('CPAYS')).Itemindex := THValComboBox(GetControl('CPAYS')).Values.IndexOf(GetParamsocsecur('SO_PAYSLOCALISATION', 'FRF'));
end;

procedure TOF_RECUPRELEVINTER.OnUpdate ;
Begin
end;


procedure TOF_RECUPRELEVINTER.OnClose ;
begin
 ListeScripts.Free;   // Liste des scripts de reconnaissance sélectionnés
 if TSV <> nil then
 begin
    TSV.TOBParamFree;
    TSV.free;
 end;
 if TobReleve <> nil then TobReleve.free;
 TCPContexte.GetCurrent.Release;
end;

procedure TOF_RECUPRELEVINTER.InitGrid ;
var i : integer ;
begin
GD.RowCount:=2;
GD.FixedRows:=1 ;
for i := 0 to NbColonne do
begin
GD.FColAligns[i]:=taCenter ;
GD.ColEditables[i] := FALSE;
end;
GD.ColEditables[1] := TRUE;
if ModeRlV <> 'EXPORT' then
begin
        GD.ColTypes[2]     := 'B';
        GD.ColFormats[2]   := IntToStr(Ord(csCheckbox));
        GD.ColTypes[3]     := 'B';
        GD.ColFormats[3]   := IntToStr(Ord(csCheckbox));
        GD.ColEditables[3] := FALSE;
        GD.ColTypes[5]     := 'B';
        GD.ColFormats[5]   := IntToStr(Ord(csCheckbox));
        GD.ColEditables[5] := FALSE;
end
else
begin
        GD.ColTypes[4]     := 'B';
        GD.ColFormats[4]   := IntToStr(Ord(csCheckbox));
        GD.ColEditables[4] := FALSE;
end;

end;


Procedure TOF_RECUPRELEVINTER.ScruteFichier (var Dossier: string; var filtre: string; Attributs: integer);
var
  FichierTrouve  : string;
  Resultat       : Integer;
  SearchRec      : TSearchRec;
  TF             : TOB;
  Sortie         : string;
begin
  if Dossier[length (Dossier)] = '\' then
    Dossier := copy (Dossier, 1, length (Dossier) -1);
  Resultat := FindFirst (Dossier + '\' + filtre, Attributs, SearchRec);
  while Resultat = 0 do
  begin
    Application.ProcessMessages;
    if ((SearchRec.Attr and faDirectory) <= 0) then // On a trouvé un Fichier (et non un dossier)
    begin
      FichierTrouve := SearchRec.Name;
    end
    else begin Resultat:= FindNext (SearchRec); continue; end;
    TF := TOB.Create('fille', TobReleve, -1);
    TF.AddChampSupValeur('Entree', FichierTrouve);
    TF.AddChampSupValeur('TableName', '');
    if ModeRlV <> 'EXPORT' then
    begin
         TF.AddChampSupValeur('Etebac', '-');
         TF.AddChampSupValeur('Releve', 'X');
    end
    else
    begin
         //sortie := ExtractFileName(FichierTrouve);
         //sortie := ChangeFileExt(FichierTrouve, '.TRA');
         sortie := 'PG'+FichierTrouve;
         TF.AddChampSupValeur('Sortie', sortie);
    end;
    TF.AddChampSupValeur('Effectue', '');
    TF.AddChampSupValeur('Variable', '');
    TF.AddChampSupValeur('Modifie', '-');
    Resultat:= FindNext (SearchRec);
  end;
  FindClose (SearchRec);// libération de la mémoire
end;

procedure TOF_RECUPRELEVINTER.BVALIDELISTEClick(Sender: TObject);
var
  TBN,Tmp           : string;
  N, index          : integer;
begin
  inherited;
  GD.refresh ;
  GD.VidePile(FALSE);

  if TobReleve <> nil then TobReleve.free;
  TobReleve := TOB.Create('Enregistrement', nil, -1);
  Repertoire := GetControlText('REPERTOIRE');
  Extent := GetControlText('EEXTENTION');
  ScruteFichier (Repertoire, Extent, FaDirectory + faHidden + faSysFile);
  if ModeRlV <> 'EXPORT' then
     TobReleve.PutGridDetail(GD, False, False, 'Entree;Tablename;Etebac;Releve;Effectue;Variables')
  else
     TobReleve.PutGridDetail(GD, False, False, 'Entree;Tablename;Sortie;Effectue;Variables');


  For N:=1 To GD.RowCount-1 do
  begin
         TBN := GD.Cells[0, N];
         if TBN = '' then continue;
         Tmp := ReadTokenPipe(TBN,'.');
         Index := Comboscript.items.IndexOf(TBN);
         if Index <> -1 then
         begin
           GD.Row := N;
           GD.ValCombo.value := TBN;
           GD.Cells[1, N]:=Comboscript.Items[Index] ;
         end;
          if ModeRlV = 'EXPORT' then
             GD.Cells[4, N] := '-'
          else
             GD.Cells[5, N] := '-'
  end;
end;

procedure TOF_RECUPRELEVINTER.BReconnaissanceClick(Sender: TObject);
var
QP : TQuery;
begin
    // Fenêtre de progression
    CallBackDlg := TCallBackDlg.Create(Application);
    // Chargement script reconnaissance
    QP := OpenSQL ('SELECT CIS_CLE from CPGZIMPREQ Where CIS_DOMAINE="R"', TRUE);
    While  not (QP.Eof) do
    begin
          if (ListeScripts.indexof(QP.FindField('CIS_CLE').asstring) < 0) then  ListeScripts.Add( QP.FindField('CIS_CLE').asstring);
          QP.Next;
    end;
    Ferme(QP);
    // Reconnaissance des scripts à appliquer aux fichiers
    RecoFichiers( Repertoire, ListeScripts, Extent, FaDirectory + faHidden + faSysFile );

    CallBackDlg.Hide;
    CallBackDlg.Free;

    if CallBackDlg.bAbort = false then PGIinfo('Reconnaissance terminée','Traitement')
    else PGIinfo('Reconnaissance annulée','Traitement');
end;

procedure TOF_RECUPRELEVINTER.GDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    32 :
      begin
          if (SavCol = 2) or  (SavCol = 3) then
          begin
            if GD.Cells[SavCol, SavRow] = 'X' then
              GD.Cells[SavCol, SavRow] := '-'
            else
              GD.Cells[SavCol, SavRow] := 'X';
          end;
      end;
  end;
end;

procedure TOF_RECUPRELEVINTER.GDSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  inherited;
  SavCol := ACol;
  SavRow := ARow;
end;

procedure TOF_RECUPRELEVINTER.GDDblClick(Sender: TObject);
var
  TBN                     : string;
  TEcr,TReleve, TF        : TOB;
begin
  inherited;
  if (SavCol = 2) or  (SavCol = 3) then
  begin
        if GD.Cells[SavCol, SavRow] = 'X' then
          GD.Cells[SavCol, SavRow] := '-'
        else
          GD.Cells[SavCol, SavRow] := 'X';
  end
  else
  begin
        if (SavCol = 0) then
        begin
            TEcr := nil; TReleve := nil;
            TBN := GD.Cells[1, SavRow];
            if TBN = '' then exit;
            Execute (TBN, TEcr, TReleve, SavRow);
            if TEcr <> nil then
            begin
                    TReleve.Detail.Sort('COMPTE;RIB;CODE');
                    LanceEtatTob('E', 'ECT', 'REL', TReleve, True, False, False, nil, '', 'Intégration de relevés bancaires', False);
                    TEcr.free;
                    TReleve.free;
            end;
        end;
        if (TobReleve <> nil) and ((SavCol = 4) or  (SavCol = 5)) then
        begin
            TF := TobReleve.FindFirst(['Entree'], [GD.Cells[0, SavRow]], FALSE);
            if TF <> nil then
            begin
                 if TF.Getvalue('Variable')= 'X' then
                 begin
                  if SaisieVariableGroupe (GD.Cells[0, SavRow], TF) then
                  begin
                       TF.putvalue('Modifie', 'X');
                       GD.Cells[SavCol, SavRow] := 'X';
                  end;
                 end
                 else GD.Cells[SavCol, SavRow] := '-';
            end;
        end;
  end;
end;

procedure TOF_RECUPRELEVINTER.ChargeTob(value : string; Row : integer);
var
FFScript       : TScript;
id,dem         : integer;
St             : string;
TF,TF1         : TOB;
begin
  inherited;

    FFScript := InitScriptGroupe(value);
    if FFscript <> nil then
    begin
          if FFScript.Variables.count = 0 then
          begin
              TF := TobReleve.FindFirst(['Entree'], [GD.Cells[0, Row]], FALSE);
              if TF <> nil then
              begin
              TF.putvalue('Tablename',GD.ValCombo.value);
              TF.putvalue('Variable', '-');
              TF.putvalue('Modifie', '-');
              GD.Cells[5, Row] := '';
              end;
          end
          else
          begin
                TF := TobReleve.FindFirst(['Entree'], [GD.Cells[0, Row]], FALSE);
                if (TF = nil) then exit;
                if (TF.Detail.count <> 0) then TF.ClearDetail;
                dem := 0;
                For id:=0 To FFScript.Variables.count-1 do
                begin
                    St := FFScript.Variables.Items[id].Libelle;
                    St := ReadTokenpipe (St,'=');
                    TF1 := TF.FindFirst(['Libelle'], [St], FALSE);
                    if TF1 = nil then
                    begin
                        TF1 := TOB.Create ('',TF,-1);
                        TF1.AddChampSupValeur('Name',FFScript.Variables.Items[id].Name);
                        TF1.AddChampSupValeur('Libelle',St);
                        TF1.AddChampSupValeur('Text', FFScript.Variables.Items[id].Text);
                        TF1.AddChampSupValeur('Demandable', FFScript.Variables.Items[id].demandable);
                    end
                    else
                    begin
                        TF1.putvalue('Name',FFScript.Variables.Items[id].Name);
                        TF1.putvalue('Libelle',St);
                        TF1.putvalue('Text', FFScript.Variables.Items[id].Text);
                    end;
                    TF.putvalue('Tablename',GD.ValCombo.value);
                    if (FFScript.Variables.Items[id].demandable) then
                    begin
                       TF.putvalue('Variable', 'X'); inc(dem);
                    end;
                    TF.putvalue('Modifie', '-');
              end;
              if dem <> 0 then
              begin
                   if ModeRlV = 'EXPORT' then
                      GD.Cells[4, Row] := 'X'
                   else
                      GD.Cells[5, Row] := 'X'
              end;
          end;
    end;
    FFScript.destroy;
end;

procedure TOF_RECUPRELEVINTER.ComboScriptChange(Sender: TObject);
begin
  inherited;
    if GD.ValCombo.value = '' then exit;
    if GD.col <> 1 then exit;
    ChargeTob(GD.ValCombo.value, GD.Row);
end;

procedure TOF_RECUPRELEVINTER.RecoFichiers(Dossier : string; Liste : TStringList; Extension : string; Attributs : integer);
var
  FichierTrouve  : string;
  SearchRec      : TSearchRec;
  i,Resultat,j   : integer;
  TF             : TOB;
begin

  if Dossier[ length(Dossier) ] = '\' then
    Dossier := copy (Dossier, 1, length(Dossier) - 1);

  Resultat := FindFirst (Dossier + '\' + Extension, Attributs, SearchRec);

  j := 1;

  while Resultat = 0 do
  begin
    if CallBackDlg.bAbort = false then           // à voir : accélerer l'annulation
    begin
      Application.ProcessMessages;
      if ((SearchRec.Attr and faDirectory) <= 0) then // On a trouvé un Fichier (et non un dossier)
      begin
        FichierTrouve := SearchRec.Name;
      end
      else
      begin
        Resultat := FindNext (SearchRec);
        continue;
      end;

      TF := TOB.Create('fille', TOBReleve, -1);
      TF.AddChampSupValeur('Entree', FichierTrouve);
      TF.AddChampSupValeur('TableName', '');
      TF.AddChampSupValeur('Etebac','-');
      TF.AddChampSupValeur('Releve', 'X');
      TF.AddChampSupValeur('Effectue', '');

      GD.Cells[0,j] := FichierTrouve;

      j := j+1;
      GD.RowCount := j;

      // Appliquons les scripts de reconnaissances sélectionnés
      // par l'utilisateur, et vérifions si les tables sont renseignées

      for i := 0 to Liste.Count-1 do
      begin
        TestFichierScripts(FichierTrouve, Liste.Strings[i]);
      end;
    end;
    Resultat:= FindNext (SearchRec);
  end;
  FindClose (SearchRec);  // Libération Mémoire
end;

function TOF_RECUPRELEVINTER.InitScriptGroupe(table : string) : TScript;
var
  S             : TmemoryStream;
  AStreamTable  : TmemoryStream;
  Ascript       : TScript;
  STable,QC     : TQuery;
begin
  inherited;
  Result := nil;
  STable := nil; S := nil; AStreamTable := nil;
  try

    STable := OpenSql ('select * from cpgzimpreq where cis_cle="'+table+'"', TRUE);
    if STable.EOF then
      begin Ferme (STable); exit; end;
    s := TmemoryStream.create;
    TBlobField(STable.FindField('CIS_PARAMETRES')).SaveToStream (s);
    s.Seek (0,0);
    QC := OpenSQl ('SELECT * FROM CPGZIMPCORRESP Where CIC_TABLEBLOB="CIS" and CIC_IDENTIFIANT="'+table+'"', TRUE);
    if not QC.EOF then
    begin
        AStreamTable := TmemoryStream.create;
        TBlobField(QC.FindField('CIC_OBJET')).SaveToStream (AStreamTable);
        AStreamTable.Seek (0,0);
    end;
    Ferme(QC);

    AScript := LoadScriptFromStream(s, AStreamTable);
  Finally
    Ferme(STable);
    s.free;
    AStreamTable.free;
  end;
  Result := Ascript;
end;

procedure TOF_RECUPRELEVINTER.ChargementTOBExp (TOBE : TOB; var Ascript : TScript);
var
  isc,Ind       : integer;
  FieldName     : string;
  TF1           : TOB;
begin
  For isc := 0 to AScript.Champ.count-1 do
  begin
    FieldName := AScript.Champ[isc].Name;
    Ind := pos('_', FieldName);
    if (Ind <> 0) then
    begin
      FieldName := copy(AScript.Champ[isc].Name, 1, Ind-1);
      AScript.Champs[isc].NomFichExt := AScript.Name+FieldName;

      TF1 := TOBE.FindFirst(['Sortief'], [AScript.Champs[isc].NomFichExt], FALSE);
      if TF1 = nil then
      begin
        TF1 := TOB.Create('', TOBE, -1);
        TF1.AddChampSupValeur('Sortief', AScript.Champs[isc].NomFichExt);
        TF1.AddChampSupValeur('FieldName', FieldName);
        TF1.AddChampSupValeur('NameScript', AScript.Name);
        TF1.AddChampSupValeur('Lieninter', '');
        TF1.AddChampSupValeur('SousEnsemble', FALSE);
        TF1.AddChampSupValeur('Orderby', '');
        TF1.AddChampSupValeur('Sortie', '');
      end;
    end;
  end;
end;

procedure TOF_RECUPRELEVINTER.TestFichierScripts( NomFichier : string; NomScript : string);
var
  FScript, AScript     : TScript;
  requete              : string;
  TOBExp, TF1          : TOB;
  i,index              : integer;
  ListeFam             : HTStringList;
  Processus            : string;
begin
  if GD.Cells[1,GD.RowCount-1] <> '' then
    exit; // Le script a déja été trouvé, passons les scripts de reco suivants

  Fscript := nil;  ListeFam := nil;
  FScript := InitScriptGroupe(NomScript);
  if Fscript <> nil then
  begin
    if Fscript.PreTrt.count <> 0 then
    begin
      // Traitement délimité
      TraiteScriptDelim(NomFichier, Repertoire + '\'+ NomFichier, Repertoire + '\'+  NomFichier+'FIXE', FALSE);
      FScript.Options.FileName := NomScript + 'FIXE';
    end;
    TOBExp := TOB.create ('Reco', nil, -1);
    ChargementTOBExp (TOBExp, FScript);


    For i:=0 To TOBExp.detail.Count-1 do
    begin
      AScript  := FScript.CloneScript;

      AScript.Assign(FScript, TOBExp.detail[i].getstring('FieldName'),'', ListeFam);

      // on force l'utilisation d'une table
      AScript.ParName := TOBExp.detail[i].getvalue('Sortief')+'.DB';

      AScript.ExecuteModeTest := TRUE;

      AScript.FEcrAna := AScript.EcrVent;
      AScript.FEcrEch := AScript.EcrEch;
      AScript.FileName := Repertoire + '\' + NomFichier;

      if (Ascript.PreTrt.count <> 0) and (FileExists(Ascript.options.fileName+'FIXE')) then
        AScript.Options.FileName := Ascript.options.fileName+'FIXE';

      AScript.ConstruitVariable(false);
      AScript.InterpreteSyntax;
      AScript.Variable.Clear;
      AjouteVariableALias(AScript.Variable);

      if AScript.Champ[0].LOrder <> '' then
      begin
        TF1 := TOBExp.FindFirst(['Sortief'], [TOBExp.detail[i].Getvalue('Sortief')], FALSE);
        if TF1 <> nil then
          TF1.PutValue('Orderby', AScript.Champ[0].LOrder);
      end;

      try AScript.Executer(CallBackDlg.LaCallback);
      except
        PGIBox('Problème avec le script ' + NomScript + ' sur le fichier ' + NomFichier,'Erreur');
        Ascript.destroy;
        CallBackDlg.Free;
        TOBExp.free;
        Fscript.destroy;
        exit;
      end;

      // La table est-elle renseignée ?  <=> le fichier est-il reconnu ?
      requete := 'SELECT * FROM ' + AScript.ParName;
      Processus := ReconnaissanceScript (Requete);
      if Processus <> '' then
      begin
        // Le fichier est reconnu par le script
        Index := Comboscript.items.IndexOf(Processus);
        if Index <> -1 then
        begin
          GD.Cells[1,GD.RowCount-1] := Comboscript.items[Index];
          GD.Col := 1;
          GD.Row := GD.RowCount-1;
          GD.ValCombo.value := Comboscript.items[Index];
        end;
      end;
      Ascript.destroy;
      FScript.destroy;

      TOBExp.free;
    end;
  end;
end;

procedure TOF_RECUPRELEVINTER.VisuFicDonnees (Sender: TObject);
var
NomCompletFicDonnees : string;
begin
  NomCompletFicDonnees := Repertoire + '\'+ GD.Cells[0, SavRow];
  if FileExists (NomCompletFicDonnees) then
    ShellExecute (Application.Handle, 'Open',  'WordPad.exe' , PChar('"' +
     NomCompletFicDonnees +'"'), PChar(NomCompletFicDonnees), SW_SHOWNORMAL)
  else
    PGIError (#13 + 'Le fichier de données' + #13 + NomCompletFicDonnees + #13 +
     'n''existe pas.' + #13 +'Visualisation impossible.', 'Echanges PGI');
end;

procedure TOF_RECUPRELEVINTER.ParamScript (Sender: TObject);
var
py : string;
begin
         TCPContexte.GetCurrent.VarCisx.CHARGECIX;
         TCPContexte.GetCurrent.VarCisx.Mode := 'I'; // mode interface appelé par la compta
         if ModeRlV = 'EXPORT' then
         begin
            Py := GetControlText('CPAYS');
            if (Py = '') then
            begin
                 PGIInfo ('Selectionnez un Pays');
                 exit;
            end;
         end;
         AglLanceFiche('CP','CPMULSCRIPT','', '', 'MODE='+ModeRlV+';PAYS='+Py)  ;
end;

procedure TOF_RECUPRELEVINTER.ParamScriptRec (Sender: TObject);
begin
         if ModeRlV = 'EXPORT' then
            ParamDictionnaire(nil, 'X', 'Relevé', 'BEL')
         else
         begin
               TCPContexte.GetCurrent.VarCisx.CHARGECIX;
               TCPContexte.GetCurrent.VarCisx.Domaine := 'R'; // reconnaissance
               TCPContexte.GetCurrent.VarCisx.Mode := 'I';   // mode interface appelé par la compta
               AglLanceFiche('CP','CPMULSCRIPT','', '', 'MODE='+ModeRlV)  ;
         end;
end;
procedure TOF_RECUPRELEVINTER.BValiderClick(Sender: TObject);
var
  TBN    : string;
  N      : integer;
  FicIni              : TIniFile;
  TEcr,TReleve        : TOB;
begin
 TEcr := nil; TReleve := nil;
  For N:=1 To GD.RowCount-1 do
  begin
         TBN := GD.Cells[1, N];
         if TBN = '' then continue;
         Execute (TBN, TEcr, TReleve, N);
  end;
  if TEcr <> nil then
  begin
             TEcr.InsertOrUpdateDB (TRUE);
             TEcr.free;
             TReleve.free;
  end;

  // sauvegarde dernier filtre utilisé
  FicIni        := TIniFile.Create(CurrentDonnee+'\GEXPCCSX.INI');
  FicIni.WriteString ('EXPORT', 'DOMAINE', GetInfoVHCX.Domaine);
  FicIni.WriteString ('EXPORT', 'ENTREE', Repertoire);
  FicIni.WriteString ('EXPORT', 'EXTENTION', Extent);

  FicIni.free;

end;

procedure TOF_RECUPRELEVINTER.BScruterClick(Sender: TObject);
var
  TBN    : string;
  N      : integer;
  FicIni              : TIniFile;
  TEcr,TReleve        : TOB;
begin
 TEcr := nil; TReleve := nil;
  For N:=1 To GD.RowCount-1 do
  begin
         TBN := GD.Cells[1, N];
         if TBN = '' then continue;
         Execute (TBN, TEcr, TReleve, N);
  end;
  if TEcr <> nil then
  begin
          TReleve.Detail.Sort('COMPTE;RIB;CODE');
          LanceEtatTob('E', 'ECT', 'REL', TReleve, True, False, False, nil, '', 'Intégration de relevés bancaires', False);
          TEcr.free;
          TReleve.free;
  end;

  // sauvegarde dernier filtre utilisé
  FicIni        := TIniFile.Create(CurrentDonnee+'\GEXPCCSX.INI');
  FicIni.WriteString ('EXPORT', 'DOMAINE', GetInfoVHCX.Domaine);
  FicIni.WriteString ('EXPORT', 'ENTREE', Repertoire);
  FicIni.WriteString ('EXPORT', 'EXTENTION', Extent);

  FicIni.free;

end;

procedure TOF_RECUPRELEVINTER.bTagClick(Sender: TObject);
var
N  : integer;
begin
  if (SavCol <> 2) and  (SavCol <> 3) then exit;
  For N:=1 To GD.RowCount-1 do
  begin
            if GD.Cells[SavCol, N] = 'X' then
              GD.Cells[SavCol, N] := '-'
            else
              GD.Cells[SavCol, N] := 'X';
  end;
end;



procedure TOF_RECUPRELEVINTER.Execute (TBN : string;  var TEcr,TReleve : TOB; N :integer);
var
  iet,I,id            : integer;
  TF                  : TOB;
  AScript,FScript     : TScript;
  NN,RepertSav        : string;
  TOBExp              : TOB;
  TF1                 : TOB;
  TableLien           : string;
  File1,File2         : string;
  CpteGeneral         : string;
  FichierIE           : TextFile;
  DatedebutExo        : TDateTime;
  DatefinExo          : TDateTime;
  Lien,FichSortie     : string;
  LLi, St             : string;
begin
  inherited;
            if ModeRlV = 'EXPORT' then
            begin
              if THValComboBox(GetControl('CPAYS')).Value = '' then
              begin
               PGIInfo ('Veuillez renseigner un pays');
               exit;
              end;
            end;
            (*chargement du script*)
            Fscript := nil;
            DatedebutExo  := iDate1900;
            DatefinExo    := iDate1900;

            FScript := InitScriptGroupe(TBN);
            if Fscript = nil then  exit;
            if Fscript.PreTrt.count <> 0 then
            begin
              // traitement délimité
              TraiteScriptDelim(TBN, Repertoire + '\'+ GD.Cells[0, N], Repertoire + '\'+  GD.Cells[0, N]+'FIXE', FALSE);
              FScript.Options.FileName   := GD.Cells[0, N]+'FIXE';
            end;
            TOBExp := TOB.create ('Export Groupe', nil, -1);
            ChargementTOBExp (TOBExp, Fscript);
            CallBackDlg := TCallBackDlg.Create(Application);
            CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
            if FScript.Variables.count <> 0 then
            begin
              TF := TobReleve.FindFirst(['Entree'], [GD.Cells[0, N]], FALSE);
              if (TF <> nil) and (TF.detail.count = FScript.Variables.count) then
              begin
                For id:=0 To TF.detail.count-1 do
                begin
                  if FScript.Variables[id].Name = TF.detail[id].GetValue('Name')then
                    FScript.Variables[id].Text := TF.detail[id].GetValue('Text');
                end;
              end;
            end;

            For iet:=0 To TOBExp.detail.Count-1 do
            begin
                          Lien := '';
                          (*création du script*)
                          AScript  := FScript.CloneScript;

                          (*affectation des champs + TFieldRec avec la fonction InitTableAndFields*)
                          AScript.Assign(FScript, TOBExp.detail[iet].getstring('FieldName'), '', nil);
                          if (AScript.FileDest = 0) then
                            AScript.ParName := ExtractFileName (TOBExp.detail[iet].getvalue('Sortief'))
                          else
                            AScript.ParName := TOBExp.detail[iet].getvalue('Sortief');

                          AScript.ExecuteModeTest := TRUE;

                          AScript.FEcrAna := AScript.EcrVent;
                          AScript.FEcrEch := AScript.EcrEch;
                          AScript.FileName := Repertoire+'\'+GD.Cells[0, N];

                          if (Ascript.PreTrt.count <> 0) and (FileExists(Ascript.options.fileName+'FIXE')) then
                            AScript.Options.FileName := Ascript.options.fileName+'FIXE';

                          if (Ascript.FileName = '') or not(FileExists(Ascript.options.fileName)) then
                          begin
                              PGiBox('Le fichier '+Ascript.options.fileName+' est manquant'+#10#13+' L''importateur ne peut continuer', 'Conception');
                              Ascript.destroy;
                              CallBackDlg.Free;
                              TOBExp.free;
                              Fscript.destroy;
                              exit;
                          end;

                          (* gestion des traitements comptables à brancher aussi sur ' E X E C I M P '*)
                          if Ascript.OptionsComptable.Correspondance.Count > 0 then
                              for I := 0 to Ascript.OptionsComptable.correspondance.count -1 do
                                  with ascript.OptionsComptable do
                                      DecoupeCritere(TInfoCpt(correspondance.Items[I]).Aux, TInfoCpt(correspondance.Items[I]).listFourchette);

                          (* Verification de la syntaxe *)
                          AScript.ConstruitVariable(false);
                          AScript.InterpreteSyntax;
                          // 0 résultat en tables paradox et 1 résultat en ASCII
                          AScript.Variable.Clear;
                          AjouteVariableALias(AScript.Variable);
                          if AScript.Champ[0].LienInter <> '' then
                          begin
                             TF1 := TOBExp.FindFirst(['Sortief'], [TOBExp.detail[iet].Getvalue('Sortief')], FALSE);
                             if TF1 <> nil then
                              TF1.PutValue('Lieninter', AScript.Champ[0].LienInter);
                              if Pos ('|', AScript.Champ[0].LienInter) = 0 then
                              begin
                                      TableLien := Copy(TOBExp.detail[iet].GetValue('Lieninter'),0, Pos('/',TOBExp.detail[iet].GetValue('Lieninter'))-1);
                                      TableLien := AScript.Name+TableLien;
                                      TF1 := TOBExp.FindFirst(['Sortief'], [TableLien], FALSE);
                                      if TF1 <> nil then
                                         TF1.PutValue('SousEnsemble', TRUE);
                              end
                              else
                              begin
                                     LLi := AScript.Champ[0].LienInter;
                                     St := ReadTokenPipe(LLi,'|');
                                     While (St <> '') do
                                     begin
                                          TableLien :=TOBExp.detail[iet].GetValue('NameScript')+Copy(St, 0, Pos('/',St)-1);
                                          TF1 := TOBExp.FindFirst(['Sortief'], [TableLien], FALSE);
                                          if TF1 <> nil then
                                             TF1.PutValue('SousEnsemble', TRUE);
                                          St := ReadTokenPipe(LLi,'|');
                                     end;
                              end;
                          end;
                          if AScript.Champ[0].LOrder <> '' then
                          begin
                             TF1 := TOBExp.FindFirst(['Sortief'], [TOBExp.detail[iet].Getvalue('Sortief')], FALSE);
                             if TF1 <> nil then
                              TF1.PutValue('Orderby', AScript.Champ[0].LOrder);
                          end;
                          try AScript.Executer(CallBackDlg.LaCallback);
                          except
                                      PGIBox('L''importation ne s''est pas correctement déroulée ','');
                                      Ascript.destroy;
                                      CallBackDlg.Free;
                                      TOBExp.free;
                                      Fscript.destroy; exit;
                          end;
                          Ascript.destroy;

            end;



            CallBackDlg.Hide;
            CallBackDlg.Free;
            CpteGeneral :='';

            if ModeRlV ='EXPORT' then
            begin
              TF := TobReleve.FindFirst(['Entree'], [GD.Cells[0, N]], FALSE);
              if TF <> nil then
              begin
               FichSortie :=  GD.Cells[2, N];
               AssignFile(FichierIE, Repertoire + '\'+ FichSortie) ;
               Rewrite(FichierIE) ;
               if IoResult<>0 then
               begin
                    PgiBox('Impossible d''écrire dans le fichier #10'+FichSortie,'Export') ;
                    CloseFile(FichierIE);
                    Exit ;
               end;
              end;
            end;

            (* Ecriture dans la base*)
            ExecuteExportFichier (FichierIE, TOBExp, TEcr, TReleve, FScript, ModeRlV, (GD.Cells[3, N]='X'), (GD.Cells[2, N]='X'), TSV, DatedebutExo, DatefinExo, THValComboBox(GetControl('CPAYS')).Value);

             if ModeRlV ='EXPORT' then CloseFile(FichierIE);
            // sauvegarde du fichier de départ dans le répertoire SAV
            RepertSav := Repertoire+'\sauvegarde\';
            if not DirectoryExists(RepertSav) then
              CreateDir(RepertSav);
            File1 := Repertoire+'\'+GD.Cells[0, N];
            File2 := RepertSav+GD.Cells[0, N];
//            Movefile (PChar(File1), PChar(File2));
            if ModeRlV ='IMPORT' then
               GD.Cells[4, N] := FormatDateTime('hh:nn:dd/mm/yyyy',NowH)
            else
               GD.Cells[3, N] := FormatDateTime('hh:nn:dd/mm/yyyy',NowH);

            For iet:=0 To TOBExp.detail.Count-1 do
              begin
                   if (FScript.FileDest = 0) then
                             NN := TOBExp.detail[iet].getvalue('Sortief')+'.DB'
                   else
                             NN := TOBExp.detail[iet].getvalue('Sortief');

                   if (FScript.FileDest = 1) then
                   begin
                               DeleteFile (CurrentDonnee+'\'+NN+'.txt');
                               DeleteFile (CurrentDonnee+'\'+NN+'.sch');
                   end
                   else
                               DeleteFile (CurrentDonnee+'\'+NN);
              end;

            if Fscript.PreTrt.count <> 0 then
               DeleteFile (Repertoire + '\' + GD.Cells[0, N]+'FIXE');
            Fscript.destroy;
            TOBExp.free;
end;

procedure TOF_RECUPRELEVINTER.RepertoireOnChange(Sender: TObject);
begin
  inherited;
  Repertoire := GetControlText('REPERTOIRE')+'\';
  Extent := GetControlText('EEXTENTION');
  if Repertoire = '' then exit;
  ChDir(Repertoire);
end;

procedure TOF_RECUPRELEVINTER.TraiteMessage (var Msg: TMsg; var Handled: Boolean);
var
  NomDuFichierStr        : string;
  NomDuFichier           : array[0..255] of char;
begin
  if Msg.message = WM_DROPFILES then
  begin
      DragQueryFile (Msg.wParam, 0, NomDuFichier, sizeof(NomDuFichier) );  // récupération du nom du fichier
      NomDuFichierStr := NomDuFichier; // tansformation du tableau de char en STRING
      DragFinish(Msg.wParam);
      if DirectoryExists (NomDuFichierStr) then
      begin
        SetWindowText (Msg.hwnd, NomDuFichier);
        BReconnaissanceClick(Application);
      end;
  end;
end;

procedure TOF_RECUPRELEVINTER.bDefaireClick(Sender: TObject);
begin
  inherited;
// RUSE pour effacer le combobox
  gd.HideEditor ;
  GD.col:=1; gd.row:=1 ;
  SendMessage(GD.Handle,WM_KEYDOWN,VK_TAB ,0) ;
  SendMessage(GD.Handle,WM_KEYDOWN,VK_left ,0) ;
  GD.refresh ;
  GD.VidePile(FALSE);
end;

function TOF_RECUPRELEVINTER.SaisieVariableGroupe (fich : string; TV : TOB): Boolean;
var
Ed1           : THCritMaskEdit;
Lb            : THLabel;
St            : string;
Lasth,Lastlh  : integer;
id            : integer;
begin
  Result := TRUE;
  FSaisieVar := TFVierge.Create(nil);
  FSaisieVar.BValider.Onclick := BValideClickVarGrp;
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
      For id:=0 To TV.detail.count-1 do
      begin
          if (not TV.detail[id].getvalue('demandable')) then continue;

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
          Ed1.Parent := FSaisieVar;
          Lasth:=Ed1.Top+Ed1.Height+2;
          Lastlh:=Ed1.Top+Ed1.Height+5;
      end;
      try
          if ShowModal <> mrOk then
               Result      := FALSE;
      finally
          Free;
      end;
      end;

end;

procedure TOF_RECUPRELEVINTER.BValideClickVarGrp(Sender: TObject);
var
id,idv    : integer;
TL        : TControl ;
TF        : TOB;
NN        : string;
begin
  inherited;
      NN := FSaisieVar.caption;
      ReadTokenpipe (NN,':');
      TF := TobReleve.FindFirst(['Entree'], [NN], FALSE);
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


Initialization
    registerclasses([TOF_RECUPRELEVINTER]);

end.
