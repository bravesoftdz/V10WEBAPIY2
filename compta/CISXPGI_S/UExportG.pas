unit UExportG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Vierge, StdCtrls, Mask, Hctrls, Grids, ComCtrls, HSysMenu, HTB97, UTOB, HPanel,
  UiUtil, uDbxDataSet, Variants, ADODB,
  db, UDMIMP, UScript, HmsgBox, Hent1, FileCtrl, Shellapi,
  PrintDBG, HXLSPAS, inifiles, URecoSelecScripts, HQRY;

type
  TFEXPGRP = class(TFVierge)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    GD: THGrid;
    ComboScript: THValComboBox;
    GroupBox1: TGroupBox;
    Label25: TLabel;
    Label1: TLabel;
    ChoixRep: THCritMaskEdit;
    Label2: TLabel;
    EExtention: TEdit;
    BValide: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    BVariables: TToolbarButton97;
    Label3: TLabel;
    RepSortie: THCritMaskEdit;
    Label12: TLabel;
    Profile: THCritMaskEdit;
    Label35: TLabel;
    EdShellExecute: THCritMaskEdit;
    bExport: TToolbarButton97;
    Reconnaissance: TToolbarButton97;
    SelectScriptsReco: TToolbarButton97;
    Domaine: THValComboBox;
    procedure BValideClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DOMAINEExit(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure GDSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GDDblClick(Sender: TObject);
    procedure BVariablesClick(Sender: TObject);
    procedure ComboScriptChange(Sender: TObject);
    procedure ChoixRepExit(Sender: TObject);
    procedure RepSortieExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BImprimerClick(Sender: TObject);
    procedure ProfileElipsisClick(Sender: TObject);
    procedure bExportClick(Sender: TObject);
    procedure SelectScriptsRecoClick(Sender: TObject);
    procedure ReconnaissanceClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBDir         : TOB;
    DomaineAppli   : string;
    FSaisieVar     : TFVierge;
    SavRow,SavCol  : Integer;
    ListeScripts   : TStringList;
    procedure ScruteFichier (var Dossier: string; var filtre: string; Attributs: integer);
    function  InitScriptGroupe(table : string) : TScript;
    procedure ChargementTOBExp (TOBE : TOB; var Ascript : TScript);
    function  SaisieVariableGroupe (fich : string; TV : TOB): Boolean;
    procedure BValideClickVarGrp(Sender: TObject);
    procedure DessineCell(ACol, ARow: Longint; Canvas: TCanvas; State:
      TGridDrawState);
    procedure TraiteMessage (var Msg: TMsg; var Handled: Boolean);
    procedure ChargeTob(value : string; Row : integer);
    procedure RecoFichiers(Dossier : string; Liste : TStringList; Extension : string; Attributs : integer);
    procedure TestFichierScripts(NomFichier : string; NomScript : string);
  public
    { Déclarations publiques }
  end;

var
  FEXPGRP: TFEXPGRP;

Procedure ExportGroupe(PP: THPanel) ;

implementation

uses
UPdomaine, UExecute, UAssistConcept, UselectCorresp;

{$R *.DFM}

Procedure ExportGroupe(PP: THPanel) ;
BEGIN
     FEXPGRP :=TFEXPGRP.Create(Application) ;
     InitInside(FEXPGRP, PP);
     FEXPGRP.Show;
     if (ParamCount > 0) then  // si lancement par la ligne de commande
        FEXPGRP.close;

END ;


Procedure TFEXPGRP.ScruteFichier (var Dossier: string; var filtre: string; Attributs: integer);
var
  FichierTrouve  : string;
  Resultat       : Integer;
  SearchRec      : TSearchRec;
  TF             : TOB;
  sortie         : string;
  Tmp,S          : string;
  F              : TextFile;
begin
  if (pos('.BAT', Uppercase(EdShellExecute.text)) <> 0) then
  begin
        AssignFile(F, EdShellExecute.text);
        {$I-} Reset (F) ; {$I+}
        if EOF(F) then begin CloseFile(F); exit;  end
        else
        begin
             Readln(F, S);
             Tmp := ReadTokenPipe(S, ' ');
        end;
        CloseFile(F);
        if pos('SCRUTE', Uppercase(Tmp)) <> 0  then
        begin
            if S = '' then exit;
            ChoixRep.text := ExtractFileDir(S);
            Dossier := ChoixRep.text;
            filtre := FusionFichierRepertoire(S, FaDirectory + faHidden + faSysFile);
        end;
  end;

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
    TF := TOB.Create('fille', TOBDir, -1);
    TF.AddChampSupValeur('Entree', FichierTrouve);
    TF.AddChampSupValeur('TableName', '');
    if DomaineAppli = 'S' then
       sortie := 'SI2'+ChangeFileExt(FichierTrouve, '.TRT')
    else
    if DomaineAppli = 'M' then
       sortie := 'IMP'+ChangeFileExt(FichierTrouve, '.SAV')
    else
       sortie := 'PG'+ChangeFileExt(FichierTrouve, '.TRA');
    sortie := UpperCase(sortie);
    TF.AddChampSupValeur('Sortie', sortie);
    TF.AddChampSupValeur('Effectue', '-');
    TF.AddChampSupValeur('DateJour','');
    TF.AddChampSupValeur('Variable', '');
    TF.AddChampSupValeur('Modifie', '-');
    Resultat:= FindNext (SearchRec);
  end;
  FindClose (SearchRec);// libération de la mémoire
end;

procedure TFEXPGRP.BValideClick(Sender: TObject);
var
  TBN,Tmp           : string;
  N, index          : integer;
  Choix,Extent      : string;
begin
  inherited;
  bDefaireClick(Sender); Index := 0;
  if TOBDir <> nil then TOBDir.free;
  TOBDir := TOB.Create('Enregistrement', nil, -1);
  Choix := ChoixRep.Text;
  Extent := EExtention.text;
  ScruteFichier (Choix, Extent, FaDirectory + faHidden + faSysFile);
  GD.ColEditables[1] := TRUE;

  GD.ColTypes[3]     := 'B';
  GD.ColFormats[3]   := IntToStr(Ord(csCheckbox));
  GD.ColEditables[3] := FALSE;
  GD.ColAligns[3]    := taCenter;

  GD.ColAligns[5] := taCenter;
  TOBDir.PutGridDetail(GD, False, False, 'Entree;Tablename;Sortie;Effectue;DateJour;Variable');
  if DomaineAppli = '' then  DomaineAppli := RendDomaine(Domaine.Text);
  ChargementComboScript (Comboscript, DomaineAppli);
  For N:=1 To GD.RowCount-1 do
  begin
         TBN := GD.Cells[0, N];
         if TBN = '' then continue;
         Tmp := ReadTokenPipe(TBN,'.');
         Index := Comboscript.items.IndexOf(TBN);
         if Index <> -1 then
         begin
           //GD.Cells[1, N] := TBN;
           GD.Row := N;
           GD.ValCombo.value := TBN;
           GD.Cells[1, N]:=Comboscript.Items[Index] ;
         end;
  end;
         TBN := GD.Cells[0, 1];
         Tmp := ReadTokenPipe(TBN,'.');
         if TBN <> '' then
         begin
           Comboscript.ItemIndex := Comboscript.items.IndexOf(TBN);
           //GD.Cells[1, 1] := TBN;
           GD.Row := 1;
           GD.Cells[1, 1]:=Comboscript.Items[Index] ;
         end;
end;

procedure TFEXPGRP.FormShow(Sender: TObject);
var
  Lib,sTmp  : string;
  FicIni    : TIniFile;
  sTempo    : string;
  Commande  : string;
  CurrentDir: string;
begin
  inherited;
  ChargementComboDomaine(Domaine, FALSE);
  // on enlève le domaine reconnaissance qui n'est pas utile ici
  Domaine.items.Delete(Domaine.Items.IndexOf('Reconnaissance'));
  Domaine.itemindex := 0;
  HMTrad.Resize(GD);
  DragAcceptFiles (ChoixRep.handle, true);
  DragAcceptFiles (RepSortie.handle, true);
  Application.OnMessage := TraiteMessage;

  // Liste des scripts de reconnaissance
  ListeScripts := TStringList.Create;

  sTempo := CurrentDonnee+'\GEXPCCSX.INI';
  if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'I') then
  begin
    Lib := RendLibelleDomaine(GetInfoVHCX.Domaine);
    Domaine.itemindex := Domaine.items.IndexOf(Lib);
    Domaine.Enabled := FALSE;
  end
  else
  begin
        if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'EXPORT') then
        begin
               Commande := ParamStr(1);
               sTmp := ReadTokenPipe (Commande, ';');
               CurrentDir := ReadTokenPipe (Commande, ';');
               if CurrentDir = '' then CurrentDir := ParamStr(1);
               if (Copy(currentDir, 1, 5) = '/INI=') then
                          sTempo := Copy(currentDir, 6, length(CurrentDir));
        end;
  end;
  if (FileExists(sTempo)) then
  begin
    // restauration dernier filtre utilisé
    FicIni              := TIniFile.Create(sTempo);
    if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'EXPORT') then
    begin
           sTmp := RendLibelleDomaine(FicIni.ReadString ('EXPORT', 'DOMAINE', ''));
           Domaine.itemindex   := Domaine.items.IndexOf(sTmp);
    end
    else
           Domaine.itemindex   := Domaine.items.IndexOf(FicIni.ReadString ('EXPORT', 'DOMAINE', ''));
    ChoixRep.text       := FicIni.ReadString ('EXPORT', 'ENTREE', '');
    EExtention.text     := FicIni.ReadString ('EXPORT', 'EXTENTION', '');
    RepSortie.text      := FicIni.ReadString ('EXPORT', 'SORTIE', '');
    EdShellExecute.Text := FicIni.ReadString ('EXPORT', 'COMMANDE', '');
    Profile.Text        := FicIni.ReadString ('EXPORT', 'PROFILE', '');
    sTmp                := FicIni.ReadString('EXPORT','RECO','');
    while ( length(sTmp) <> 0 ) do ListeScripts.add(ReadTokenPipe(sTmp,';'));   // restauration liste scripts reco

    FicIni.free;
  end
  else
  // on positionne par défaut le combo de domaine sur "Comptabilité PGI"
  Domaine.itemindex := Domaine.Items.IndexOf('Comptabilité PGI');
  if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'EXPORT') and (GetInfoVHCX.Domaine = 'R') then
  begin
     ReconnaissanceClick(Sender);
     
  end;
end;

procedure TFEXPGRP.DOMAINEExit(Sender: TObject);
begin
  inherited;
  DomaineAppli := RendDomaine(Domaine.Text);
  ChargementComboScript (Comboscript, DomaineAppli);
end;

procedure TFEXPGRP.bDefaireClick(Sender: TObject);
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

function TFEXPGRP.InitScriptGroupe(table : string) : TScript;
var
  STable        : TADOTable;
  S             : TmemoryStream;
  AStreamTable  : TmemoryStream;
  Ascript       : TScript;
begin
  inherited;
  Result := nil;
  STable := nil; S := nil; AStreamTable := nil;
  try
    STable := TADOTable.Create(Application);
    STable.Connection := DMImport.DBGLOBAL;
    STable := DMImport.GzImpReq;
    STable.Open;
    AStreamTable := TmemoryStream.create;
    if not STable.Locate('CLE', table, [loCaseInsensitive]) then
       exit;
    s := TmemoryStream.create;
    TBlobField(STable.FieldByName('PARAMETRES')).SaveToStream (s);
    s.Seek (0,0);
    TBlobField(STable.FieldByName('TBLCOR')).SaveToStream (AStreamTable);
    AStreamTable.Seek (0,0);

    AScript := LoadScriptFromStream(s, AStreamTable);
  Finally
    STable.Close;
    s.free;
    AStreamTable.free;
  end;
  Result := Ascript;
end;

procedure TFEXPGRP.ChargementTOBExp (TOBE : TOB; var Ascript : TScript);
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
      end;
    end;
  end;
end;

procedure TFEXPGRP.BValiderClick(Sender: TObject);
var
  N,iet,I,id          : integer;
  TF                  : TOB;
  AScript,FScript     : TScript;
  TBN,NN,RepertSav    : string;
  TOBExp              : TOB;
  FichierIE           : TextFile;
  DatedebutExo, DatefinExo : TDateTime;
  TF1                 : TOB;
  TableLien           : string;
  File1,File2         : string;
  ListeFam            : HTStringList;
  FicIni              : TIniFile;

  procedure CreateListeFam;
  var
    Q1    : TQuery;
  begin
    if Profile.Text='' then exit;
    ListeFam := HTStringList.create;
    if  not DMImport.DBGlobal.Connected then
    DMImport.DBGlobal.Connected := TRUE ;
    Q1 := OpenSQLADO ('SELECT ChampCode from '+DMImport.GzImpCorresp.TableName+ ' Where Profile="'+ Profile.text+'"', DMImport.DBGlobal.ConnectionString);
    while not Q1.EOF do
    begin
      ListeFam.add(Q1.FindField ('ChampCode').asstring);
      Q1.next;
    end;
     Q1.close;
  end;

begin
  inherited;
  CreateListeFam;
  For N:=1 To GD.RowCount-1 do
  begin
         TBN := GD.Cells[1, N];
         if TBN = '' then continue;

            (*chargement du script*)
            Fscript := nil;
            FScript := InitScriptGroupe(TBN);
            if Fscript = nil then  continue;
            if Fscript.PreTrt.count <> 0 then
            begin
              // traitement délimité
              TraiteScriptDelim(TBN, ChoixRep.Text + '\'+ GD.Cells[0, N], ChoixRep.Text + '\'+  GD.Cells[0, N]+'FIXE', FALSE);
              FScript.Options.FileName   := GD.Cells[0, N]+'FIXE';
            end;
            TOBExp := TOB.create ('Export Groupe', nil, -1);
            ChargementTOBExp (TOBExp, Fscript);
            CallBackDlg := TCallBackDlg.Create(Self);
            CallbackDlg.Label1.Caption := 'Vérification de la syntaxe...';
            if FScript.Variables.count <> 0 then
            begin
              TF := TOBDir.FindFirst(['Entree'], [GD.Cells[0, N]], FALSE);
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
                          (*création du script*)
                          AScript  := FScript.CloneScript;

                          (*affectation des champs + TFieldRec avec la fonction InitTableAndFields*)
                          AScript.Assign(FScript, TOBExp.detail[iet].getstring('FieldName'), Profile.text, ListeFam);
                          AScript.ParName := TOBExp.detail[iet].getvalue('Sortief');

                          AScript.ExecuteModeTest := TRUE;

                          AScript.FEcrAna := AScript.EcrVent;
                          AScript.FEcrEch := AScript.EcrEch;
                          AScript.FileName := ChoixRep.Text+'\'+GD.Cells[0, N];

                          if (Ascript.PreTrt.count <> 0) and (FileExists(Ascript.options.fileName+'FIXE')) then
                            AScript.Options.FileName := Ascript.options.fileName+'FIXE';

                          if (Ascript.FileName = '') or not(FileExists(Ascript.options.fileName)) then
                          begin
                              PGiBox('Le fichier '+Ascript.options.fileName+' est manquant'+#10#13+' L''importateur ne peut continuer', 'Conception');
                              Ascript.destroy;
                              CallBackDlg.Free;
                              TOBExp.free;
                              Fscript.destroy;
                              if Profile.Text<>'' then ListeFam.free;
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
                              TableLien := Copy(TOBExp.detail[iet].GetValue('Lieninter'),0, Pos('/',TOBExp.detail[iet].GetValue('Lieninter'))-1);
                              TableLien := AScript.Name+TableLien;
                              TF1 := TOBExp.FindFirst(['Sortief'], [TableLien], FALSE);
                              if TF1 <> nil then
                                 TF1.PutValue('SousEnsemble', TRUE);
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
            GD.Cells[3, N] := 'X';     // on coche la case "Effectué"
            TBN := GD.Cells[2, N];
            AssignFile(FichierIE,  RepSortie.text+'\'+TBN) ;
            Rewrite(FichierIE);
            if IoResult<>0 then
            begin
                              PgiBox('Impossible d''écrire dans le fichier #10'+TBN,'Export') ;
                              CloseFile(FichierIE);
                              Exit ;
            end ;

            (* Ecriture du fichier *)
            For iet:=0 To TOBExp.detail.Count-1 do
            begin
              Application.ProcessMessages;
                          if TOBExp.detail[iet].GetValue('SousEnsemble') then continue;

                          NN := TOBExp.detail[iet].getvalue('Sortief');
                          if DomaineAppli = '' then  DomaineAppli := RendDomaine(Domaine.Text);
                          if (TOBExp.detail[iet].GetValue('Lieninter') <> '') and (FScript.FileDest = 0) then
                               TableLien :=TOBExp.detail[iet].GetValue('NameScript')+Copy(TOBExp.detail[iet].GetValue('Lieninter'), 0, Pos('/',TOBExp.detail[iet].GetValue('Lieninter'))-1)
                          else
                          TableLien := '';

                          if (FScript.FileDest = 0) then
                             EcritureFichierExport(FichierIE, NN, TOBExp.detail[iet].getvalue('FieldName'), DomaineAppli, DatedebutExo, DatefinExo, TOBExp.detail[iet].GetValue('Lieninter'), '',TableLien, TOBExp.detail[iet].GetValue('Orderby'))
                          else
                             ExportFichierText (FichierIE, CurrentDonnee+'\'+NN+'.txt');
            end;
            CloseFile(FichierIE);

            // sauvegarde du fichier de départ dans le répertoire SAV
            RepertSav := ChoixRep.Text+'\sauvegarde\';
            if not DirectoryExists(RepertSav) then
              CreateDir(RepertSav);
            File1 := ChoixRep.Text+'\'+GD.Cells[0, N];
            File2 := RepertSav+GD.Cells[0, N];
            Movefile (PChar(File1), PChar(File2));
            GD.Cells[4, N] := FormatDateTime('hh:nn:dd/mm/yyyy',NowH);

            For iet:=0 To TOBExp.detail.Count-1 do
              begin
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
               DeleteFile (ChoixRep.Text + '\' + GD.Cells[0, N]+'FIXE');
            Fscript.destroy;
            TOBExp.free;
  end;
  if Profile.Text<>'' then ListeFam.free;
  // sauvegarde dernier filtre utilisé
  FicIni        := TIniFile.Create(CurrentDonnee+'\GEXPCCSX.INI');
  FicIni.WriteString ('EXPORT', 'DOMAINE', Domaine.Text);
  FicIni.WriteString ('EXPORT', 'ENTREE', ChoixRep.text);
  FicIni.WriteString ('EXPORT', 'SORTIE', RepSortie.Text);
  FicIni.WriteString ('EXPORT', 'COMMANDE', EdShellExecute.text);
  FicIni.WriteString ('EXPORT', 'PROFILE', PROFILE.TEXT);
  FicIni.WriteString ('EXPORT', 'EXTENTION', EExtention.TEXT);

  FicIni.free;

end;


function TFEXPGRP.SaisieVariableGroupe (fich : string; TV : TOB): Boolean;
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
          if ShowModal <> mrOk then
               Result      := FALSE;
      finally
          Free;
      end;
      end;

end;


procedure TFEXPGRP.BValideClickVarGrp(Sender: TObject);
var
id,idv    : integer;
TL        : TControl ;
TF        : TOB;
NN        : string;
begin
  inherited;
      NN := FSaisieVar.caption;
      ReadTokenpipe (NN,':');
      TF := TOBDir.FindFirst(['Entree'], [NN], FALSE);
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

procedure TFEXPGRP.GDSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  inherited;
  SavCol := ACol;
  SavRow := ARow;
end;

procedure TFEXPGRP.GDDblClick(Sender: TObject);
var
TF        : TOB;
begin
  inherited;
  TF := TOBDir.FindFirst(['Entree'], [GD.Cells[0, SavRow]], FALSE);
  if TF <> nil then
  begin
       if TF.Getvalue('Variable')= 'X' then
       begin
        if SaisieVariableGroupe (GD.Cells[0, SavRow], TF) then
        begin
             TF.putvalue('Modifie', 'X');
             GD.Cells[5, SavRow] := '&';
        end;
       end
       else GD.Cells[5, SavRow] := '';
  end;
end;

procedure TFEXPGRP.BVariablesClick(Sender: TObject);
var
N  : integer;
TF : TOB;
begin
  inherited;
  For N:=1 To GD.RowCount-1 do
  begin
        TF := TOBDir.FindFirst(['Entree'], [GD.Cells[0, N]], FALSE);
        if TF <> nil then
        begin
             if TF.Getvalue('Variable')= 'X' then
             begin
              SaisieVariableGroupe (GD.Cells[0, N], TF);
              GD.Cells[5, N] := '&';
             end;
        end;
  end;
end;


procedure TFEXPGRP.DessineCell(ACol, ARow: Longint; Canvas: TCanvas; State:
  TGridDrawState);
var
  Texte: string;
  Rect: TRect;
begin
  Rect := Gd.CellRect(ACol, ARow);
  // si autre que la colonne des coches : on ne fait rien
  if (ARow = 0) or(ACol <> 5) then
  exit;
  // case à cocher (marquage)
  (*if Gd.Cells[ACol, Arow] = 'X' then
    Texte := 'þ' // coche
  else
  *)
  if (Gd.Cells[ACol, Arow] = 'X') or (Gd.Cells[ACol, Arow] = '&') then
    Texte := '&' // book
  else
    Texte := '';
    //Texte := '¨'; // case vide
  with Gd.Canvas do
  begin
    FillRect(Rect); // redessine
    Font.Name := 'Wingdings';
    Font.Size := 10;
    if Gd.Cells[ACol, Arow] <> '&' then Font.Color := clOlive
    else
    Font.Color := clWindowText;
    // centrage de la coche
    TextOut((Rect.Left + Rect.Right) div 2 - TextWidth(Texte) div 2,
      (Rect.Top + Rect.Bottom) div 2 - TextHeight(Texte) div 2, Texte);
  end;

end;

procedure TFEXPGRP.ChargeTob(value : string; Row : integer);
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
              TF := TOBDir.FindFirst(['Entree'], [GD.Cells[0, Row]], FALSE);
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
                TF := TOBDir.FindFirst(['Entree'], [GD.Cells[0, Row]], FALSE);
                if (TF = nil) then exit;
                if (TF.Detail.count <> 0) then TF.ClearDetail;
                dem := 0;
                For id:=0 To FFScript.Variables.count-1 do
                begin
//                     if (not FFScript.Variables.Items[id].demandable) then continue;

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
              if dem <> 0 then GD.Cells[5, Row] := 'X';
          end;
    end;
    FFScript.destroy;
// TOBDir.SaveToFile('XXXX.tmp',True,True,True);
end;


procedure TFEXPGRP.ComboScriptChange(Sender: TObject);
begin
  inherited;
    if GD.ValCombo.value = '' then exit;
    if GD.col <> 1 then exit;
    ChargeTob(GD.ValCombo.value, GD.Row);
end;

procedure TFEXPGRP.ChoixRepExit(Sender: TObject);
begin
  inherited;
  if ChoixRep.Text = '' then exit;
    if not DirectoryExists(ChoixRep.Text) then
    begin
       if (PGIAsk('Chemin inexistant, voulez-vous le créer : ?', 'CISX')= mryes) then
         CreateDir(ChoixRep.Text)
       else exit;
    end;
    ChDir(ChoixRep.Text);
end;

procedure TFEXPGRP.RepSortieExit(Sender: TObject);
begin
  inherited;
  if RepSortie.Text = '' then exit;
    if not DirectoryExists(RepSortie.Text) then
    begin
       if (PGIAsk('Chemin '+RepSortie.Text+' inexistant, voulez-vous le créer : ?', 'CISX')= mryes) then
       CreateDir(RepSortie.Text)
       else exit;
    end;
end;


procedure TFEXPGRP.TraiteMessage (var Msg: TMsg; var Handled: Boolean);
var
  NomDuFichierStr        : string;
  NomDuFichier           : array[0..255] of char;
begin
  if Msg.message = WM_DROPFILES then
  begin
    if (Msg.hwnd = ChoixRep.handle) or  (Msg.hwnd = RepSortie.handle)then
    begin
      DragQueryFile (Msg.wParam, 0, NomDuFichier, sizeof(NomDuFichier) );  // récupération du nom du fichier
      NomDuFichierStr := NomDuFichier; // tansformation du tableau de char en STRING
      DragFinish(Msg.wParam);
      if DirectoryExists (NomDuFichierStr) then
      begin
        SetWindowText (Msg.hwnd, NomDuFichier);
        if (Msg.hwnd = ChoixRep.handle) then BValideClick(Application);
      end;
    end;
  end;
end;



procedure TFEXPGRP.FormCreate(Sender: TObject);
begin
  inherited;
  GD.PostDrawCell := DessineCell;
end;

procedure TFEXPGRP.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if TOBDir <> nil then TOBDir.free;
  GD.VidePile(FALSE);
  ListeScripts.Free;   // Liste des scripts de reconnaissance sélectionnés
end;

procedure TFEXPGRP.BImprimerClick(Sender: TObject);
begin
  inherited;
  PrintDBGrid (GD,Nil, Caption,'','');
end;



procedure TFEXPGRP.ProfileElipsisClick(Sender: TObject);
begin
  inherited;
  Profile.Text := SelectCorresp('Liste de correspondance');
end;

procedure TFEXPGRP.bExportClick(Sender: TObject);
begin
  inherited;
  ExportGrid(GD, nil, 'GridExport.txt', 0, TRUE);
end;

procedure TFEXPGRP.SelectScriptsRecoClick(Sender: TObject);
var
  SelScript    : TFSelecScript;
  i            : integer;
  sTmp         : string;
  FicIni       : TIniFile;
begin
  inherited;
  Selscript := nil;
  try
    SelScript := TFSelecScript.create(self);

    // Chargement scripts déja sélectionnés +  suppression de la liste scripts dispos
    if ListeScripts.Text <> '' then
    begin
      for i:=0 to ListeScripts.Count-1 do
      begin
        SelScript.ListScriptsDispo.Items.Delete( SelScript.ListScriptsDispo.Items.IndexOf(ListeScripts.Strings[i]) );
        SelScript.ListScriptsSelec.Items.Add( ListeScripts.Strings[i] );
      end;
    end;

    SelScript.showmodal;

  finally
    if Selscript.ModalResult = 1 then
    begin
      // Si utilisateur a validé
      if Selscript.ListScriptsSelec.Items.text <> '' then
      begin
        // Si un ou plusieurs scripts sont sélectionnés
        if ListeScripts.Text <> '' then ListeScripts.Clear;

        for i := 0 to SelScript.ListScriptsSelec.Items.count-1 do
        begin
          ListeScripts.Add( SelScript.ListScriptsSelec.items[i] );
        end;

        sTmp := '';

        if (FileExists(CurrentDonnee+'\GEXPCCSX.INI')) then
        begin
          // Enregistrement de la liste des scripts dans .ini
          FicIni := TIniFile.Create(CurrentDonnee+'\GEXPCCSX.INI');
          for i := 0 to ListeScripts.Count-1 do
          begin
            sTmp := sTmp + ListeScripts.Strings[i] + ';' ;
          end;
          FicIni.WriteString('EXPORT','RECO',sTmp);

          FicIni.Free;
        end;
      end;
    end;
    SelScript.Free;
  end;
end;

procedure TFEXPGRP.RecoFichiers(Dossier : string; Liste : TStringList; Extension : string; Attributs : integer);
var
  FichierTrouve  : string;
  S,sTmp         : string;
  sortie         : string;
  F              : textfile;
  SearchRec      : TSearchRec;
  i,Resultat,j   : integer;
  TF             : TOB;
begin
  if (pos('.BAT', Uppercase(EdShellExecute.text)) <> 0) then
  begin
    AssignFile(F, EdShellExecute.text);
    {$I-} Reset (F) ; {$I+}
    if EOF(F) then begin CloseFile(F); exit;  end
    else
    begin
      Readln(F, S);
      sTmp := ReadTokenPipe(S, ' ');
    end;

    CloseFile(F);

    if pos('SCRUTE', Uppercase(sTmp)) <> 0  then
    begin
      if S = '' then exit;
      ChoixRep.text := ExtractFileDir(S);
      Dossier   := ChoixRep.text;
      Extension := FusionFichierRepertoire(S, FaDirectory + faHidden + faSysFile);
    end;
  end;

  if Dossier[ length(Dossier) ] = '\' then
    Dossier := copy (Dossier, 1, length(Dossier) - 1);

  Resultat := FindFirst (Dossier + '\' + Extension, Attributs, SearchRec);

  j := 1;
  ChargementComboScript (Comboscript, '');  // '' pour avoir tous les scripts dans le combo

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

      TF := TOB.Create('fille', TOBDir, -1);
      TF.AddChampSupValeur('Entree', FichierTrouve);
      TF.AddChampSupValeur('TableName', '');
      if DomaineAppli = 'S' then
        sortie := 'SI2'+ChangeFileExt(FichierTrouve, '.TRT')
      else
      if DomaineAppli = 'M' then
        sortie := 'IMP'+ChangeFileExt(FichierTrouve, '.SAV')
      else
        sortie := 'PG'+ChangeFileExt(FichierTrouve, '.TRA');
      sortie := UpperCase(sortie);
      TF.AddChampSupValeur('Sortie', sortie);
      TF.AddChampSupValeur('Effectue', '-');
      TF.AddChampSupValeur('DateJour','');
      TF.AddChampSupValeur('Variable', '');
      TF.AddChampSupValeur('Modifie', '-');

      GD.Cells[0,j] := FichierTrouve;
      sTmp := FichierTrouve;
      GD.Cells[2,j] := sortie;
      GD.Cells[3,j] := '-';
      GD.Cells[4,j] := '';
      GD.Cells[5,j] := '';
    //GD.Cells[6,j] := '-';

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
  if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'EXPORT') then
      TOBDir.SaveToFile(RepSortie.text,True,True,True);
end;

procedure TFEXPGRP.TestFichierScripts( NomFichier : string; NomScript : string);
var
  FScript, AScript     : TScript;
  requete              : string;
  TOBExp, TF1          : TOB;
  i,index              : integer;
  Q                    : THQuery;
  ListeFam             : HTStringList;
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
      TraiteScriptDelim(NomFichier, ChoixRep.Text + '\'+ NomFichier, ChoixRep.Text + '\'+  NomFichier+'FIXE', FALSE);
      FScript.Options.FileName := NomScript + 'FIXE';
    end;
    TOBExp := TOB.create ('Reco', nil, -1);
    ChargementTOBExp (TOBExp, FScript);


    For i:=0 To TOBExp.detail.Count-1 do
    begin
      AScript  := FScript.CloneScript;

      AScript.Assign(FScript, TOBExp.detail[i].getstring('FieldName'),Profile.text, ListeFam);

      // on force l'utilisation d'une table
      AScript.ParName := TOBExp.detail[i].getvalue('Sortief');

      AScript.ExecuteModeTest := TRUE;

      AScript.FEcrAna := AScript.EcrVent;
      AScript.FEcrEch := AScript.EcrEch;
      AScript.FileName := ChoixRep.Text + '\' + NomFichier;

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

      Q := THQuery.Create(nil);
      Q.ConnectionString := DMImport.DBGlobalD.ConnectionString;
      Q.SQL.Add (requete);
      Q.Open;

      if Q.FindField('Origine_Editeur').AsString <> '' then
      begin
        // Le fichier est reconnu par le script
        Index := Comboscript.items.IndexOf(Q.FindField('Origine_Processus').AsString);
        if Index <> -1 then
        begin
          GD.Cells[1,GD.RowCount-1] := Comboscript.items[Index];
          GD.Col := 1;
          GD.Row := GD.RowCount-1;
          GD.ValCombo.value := Comboscript.items[Index];
        end;
      end;

      Q.Close;
      Q.Free;

      Ascript.destroy;
      FScript.destroy;

      TOBExp.free;
    end;
  end;
end;

procedure TFEXPGRP.ReconnaissanceClick(Sender: TObject);
begin
  inherited;
  if ListeScripts.Count <> 0 then
  begin
    if TOBDir <> nil then TOBDir.free;
    TOBDir := TOB.Create('Enregistrement', nil, -1);

    GD.ColEditables[1] := TRUE;
    GD.ColTypes[3]     := 'B';
    GD.ColFormats[3]   := IntToStr(Ord(csCheckbox));
    GD.ColEditables[3] := FALSE;
    GD.ColAligns[3]    := taCenter;
    GD.ColAligns[5]    := taCenter;
    TOBDir.PutGridDetail(GD, False, False, 'Entree;Tablename;Sortie;Effectue;DateJour;Variable');
    // Fenêtre de progression
    CallBackDlg := TCallBackDlg.Create(Self);

    // Reconnaissance des scripts à appliquer aux fichiers
    RecoFichiers( ChoixRep.text, ListeScripts, EExtention.text, FaDirectory + faHidden + faSysFile );

    CallBackDlg.Hide;
    CallBackDlg.Free;

    if CallBackDlg.bAbort = false then PGIinfo('Reconnaissance terminée','Traitement')
    else PGIinfo('Reconnaissance annulée','Traitement');
  end
  else
    PGIinfo('Vous n''avez pas sélectionné de scripts de reconnaissance','Attention');
end;

end.
