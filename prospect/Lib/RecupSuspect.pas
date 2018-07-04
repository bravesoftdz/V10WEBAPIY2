unit RecupSuspect;

interface
uses
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  Hent1, UIUtil, HPanel, UTOB, KPMGUtil,utobdebug,ent1,Aglinit,
  {$IFNDEF EAGLCLIENT}
    Tablette,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ELSE}
    etablette,
  {$ENDIF}
  ed_tools, Grids,TrtRecupSuspect ,M3FP,Paramsoc,UtilRT,UtilPgi,EntGC, Mask;

procedure EntreeRecupSuspect (StParSuspect : string);
procedure SupRecupSuspect (StParSuspect : string);

Const
  SAF_Libelle           : integer = 0;
  SAF_Offset            : integer = 1;
  SAF_Longueur          : integer = 2;
  SAF_Champ             : integer = 3;
  SAF_Champ2            : integer = 4;
  SAF_Champ3            : integer = 5;
  KBlock                : integer = 1000;     // taille des blocks de lignes suspects à intégrer
  KBlockTiers           : integer = 200;      // taille des blocks pour maj tiers


type
  TFRecupSuspect = class(TFAssist)
    TINTRO: THLabel;
    Fichier: TTabSheet;
    Criteres: TTabSheet;
    Recapitulatif: TTabSheet;
    PTITRE: THPanel;
    HLabel1: THLabel;
    GBCreationFiches: TGroupBox;
    CBCreationSuspect: TCheckBox;
    TNomChamp: THLabel;
    TLongueur: THLabel;
    TOffset: THLabel;
    TCritere: THLabel;
    TBorneSup: THLabel;
    TBorneInf: THLabel;
    TTextCritere: THLabel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    ListRecap: TListBox;
    GBPremier: TGroupBox;
    HChamp1: THValComboBox;
    HLongueur1: THCritMaskEdit;
    HOffset1: THCritMaskEdit;
    HBorneInf1: THCritMaskEdit;
    HBorneSup1: THCritMaskEdit;
    GBDeuxieme: TGroupBox;
    HChamp2: THValComboBox;
    HLongueur2: THCritMaskEdit;
    HOffset2: THCritMaskEdit;
    HBorneInf2: THCritMaskEdit;
    HBorneSup2: THCritMaskEdit;
    GBTroisieme: TGroupBox;
    HChamp3: THValComboBox;
    HLongueur3: THCritMaskEdit;
    HOffset3: THCritMaskEdit;
    HBorneInf3: THCritMaskEdit;
    HBorneSup3: THCritMaskEdit;
    TRSS_PARSUSPECT: THLabel;
    RSS_PARSUSPECT: THCritMaskEdit;
    HRecap: THMsgBox;
    OpenDialogButton: TOpenDialog;
    TRSS_FICHIER: THLabel;
    RSS_FICHIER: THCritMaskEdit;
    Apercu: TTabSheet;
    GChamp: THGrid;
    CBControleProspect: TCheckBox;
    LRSS_PARSUSPECT: TLabel;
    HLabel3: THLabel;
    OrigineFichier: THValComboBox;
    CBPremiereLigne: TCheckBox;
    bparametre: TToolbarButton97;
    CBMajTiers: TCheckBox;
    GBTiers: TGroupBox;
    CBtest: TCheckBox;
    CBContact: TCheckBox;
    CbProspect: TCheckBox;
    CbClient: TCheckBox;
    CbAutre: TCheckBox;
    procedure bFinClick(Sender: TObject);
    procedure RSS_FICHIERElipsisClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure HChamp1Change(Sender: TObject);
    procedure HChamp2Change(Sender: TObject);
    procedure HChamp3Change(Sender: TObject);
    procedure RSS_PARSUSPECTExit(Sender: TObject);
    procedure BDescriptionClick(Sender: TObject);
    procedure Formshow(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    function  NextPage : TTabSheet ; Override ;
    procedure CBControleProspectClick(Sender: TObject);
    procedure CBMajTiersClick(Sender: TObject);
    procedure CBtestClick(Sender: TObject);
//    procedure BLigneSClick(Sender: TObject);
  private
    { Déclarations privées }
    TobSuspect          : TOB;
    TobSuspectCompl     : TOB;
    stEnreg             : string;
    BSTop               : Boolean;
    passage             : Boolean;
    LongFixe            : Boolean;
    Particulier         : Boolean;
    Nb_Ecrit            : integer;
    Nb_ErreurSiret      : integer;
//    SQL_Insert, SQL_Update, SQL_Champs, SQL_Datas, SQL_Where, SQL_Sep,
    compteurCode        : string;
    stAucun,stAutre     : string;
    BoAjoutCorr         : Boolean;
    bMajCodeTiers       : Boolean;
    NumchampCodeTiers   : integer;
    TobTiers            : Tob;
    TobYtc              : Tob;
    TobRpr              : Tob;

    // Initialisations
    procedure ActiveChamp (HCrit : THCritMaskEdit);
    procedure DepileTOBLigne ;
    procedure DesactiveChamp (HCrit : THCritMaskEdit);
    procedure InitialiseFiche;
    procedure RempliCombo;
    // Controles
    function ControleChampOk (HChamp : THValComboBox; ValChamp : string;
                              stLongueur, stOffset, stBorneInf, stBorneSup : string) : boolean;
    function ControleEnreg : boolean;
    function controleSaisie : boolean;
    function ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur,
                                stOffset : string; Message1, Message2 : integer) : boolean;
    // Traitements
    procedure CompleteChampSuspect;
    procedure ListeRecap;
    procedure RecapChamp(stChamp, stLongueur, stOffset, stBorneInf, stBorneSup : string);
    procedure RecupereSuspect;
    procedure RecupereValChamp (TobPar : TOB; Index : integer; var ValChamp : string);
    procedure RecupValChampEnregVar (var ValChamp : string; IndexChamp : integer);
    procedure RenseigneJournalEvenement (ioerr : TIOErr);
    procedure TraiteChamp (IndexTobTable : integer; ValChamp, Champ : string);
    procedure TraiteEnregistrement (Apercu:boolean) ;
    // Utils
    function  ExtractLibelle ( St : string) : string;
//    procedure ImportEnregSuspect;
//    procedure ImportEnregSuspectCompl;
    procedure EnregLesSuspects;
//    procedure Ecrire_Datas(NomChampTable : string; NumTable, NumeroChamp : integer;
//                     ValeurSeq : variant; var SQL_Champs, SQL_Datas, SQL_Sep : string);
    Procedure InitialiseApercu;
    procedure TraiteEnregistrementApercu ( Ligne : string);
    function  AffectCodeSuspect: string;
    function  InitProgression: longint;
    procedure CorrespondanceJur;
    procedure CorrespondanceFonc ;
    procedure CorrespondanceEff;
    procedure CorrespondanceCA;
    procedure CorrespondanceDateCreation;
    function  VerifAccesFichier : boolean;
    procedure GereCombo (Champ, ValChamp : String; Latob : TOB);
    procedure ChargeTablette;
    procedure ChargeTableCorresp;
    procedure ChercheProchainCode;
    Function TraduireFormule (LaFormule, LeChamp : String) : String;
    procedure MajTiersParCode;
    function ExisteTiers (CodeTiers : String) : boolean;
    procedure TraiteChampCodeTiers (LeChamp : String; LeValChamp : String);
    procedure EnregLesTiers;
    procedure ListRecapMajCodeTiers;


  public
    { Déclarations publiques }
    TobParSuspect       : TOB;
    TobParSuspectLig    : TOB;
    TobCorresp          : TOB;
    TobTablette         : TOB;
    ProchainCodeLibre   : String;


  end;

var
  FRecupSuspect: TFRecupSuspect;

implementation

uses TntForms,UFonctionsCBP,
 		CbpMCD
   ,CbpEnumerator;
   
{$R *.DFM}

procedure EntreeRecupSuspect (StParSuspect : string);
var
  Fo_Assist             : TFRecupSuspect;
  PPANEL                : THPanel ;
BEGIN
  Fo_Assist             := TFRecupSuspect.Create(Application) ;
  Fo_Assist.RSS_PARSUSPECT.text := StParSuspect;
  PPANEL                := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL

  if PPANEL = Nil then
  begin
    try
      Fo_Assist.ShowModal ;
    finally
      Fo_Assist.Free ;
    end ;
    end else
    BEGIN
      InitInside (Fo_Assist, PPANEL);
      Fo_Assist.Show ;
    end ;
end ;

procedure SupRecupSuspect (StParSuspect : string);
Begin
  if (trim (StParSuspect) = '') then exit;
  if PGIAsk('Confirmez-vous la suppression du paramétrage de l''import des suspects?','Code Import Suspects :'+StParSuspect)<>mrYes then exit ;
  BeginTrans;
  ExecuteSQL('DELETE FROM PARSUSPECTLIG WHERE RRL_PARSUSPECT="' +StParSuspect+ '"');
  ExecuteSQL('DELETE FROM PARSUSPECT WHERE RSS_PARSUSPECT="' +StParSuspect+ '"');
  try
      CommitTrans;
  except
      RollBack;
  end;
end;

procedure TFRecupSuspect.FormCreate(Sender: TObject);
begin
  inherited;
TobParSuspect := TOB.Create ('PARSUSPECT', Nil, -1);
TobParSuspectLig := TOB.Create ('', Nil, -1);
TobSuspect := TOB.Create ('SUSPECTS', Nil, -1);
TobSuspect.InitValeurs;
TobSuspectCompl := TOB.Create ('SUSPECTSCOMPL', Nil, -1);
TobSuspectCompl.InitValeurs;
stAucun := TraduireMemoire('Aucun');
stAutre := TraduireMemoire('Autre');
end;

procedure TFRecupSuspect.FormShow(Sender: TObject);
begin
  inherited;
  bMajCodeTiers         := False;
  initialisefiche;
end;
{==============================================================================================}
{=============================== Evènements de la Forme========================================}
{==============================================================================================}


procedure TFRecupSuspect.RSS_PARSUSPECTExit(Sender: TObject);
begin
  inherited;
 if RSS_PARSUSPECT.text <> '' then InitialiseFiche;
end;

procedure TFRecupSuspect.RSS_FICHIERElipsisClick(Sender: TObject);
begin
  inherited;
if OpenDialogButton.Execute then
    if OpenDialogButton.FileName <> '' then
        RSS_FICHIER.Text := OpenDialogButton.Filename;
end;

procedure TFRecupSuspect.BDescriptionClick(Sender: TObject);
var Action : TActionFiche;
begin
  inherited;
  if VerifAccesFichier = True then
  begin
    if (RSS_PARSUSPECT.Text <> '') then  action := taModif
    else action := taCreat;
    RSS_PARSUSPECT.Text := EntreeTrtRecupSuspect(action,RSS_PARSUSPECT.Text,RSS_FICHIER.Text);
    InitialiseFiche;
    if LoadLesTobParSuspect (RSS_PARSUSPECT.Text, TobParSuspect, TobParSuspectLig) then
    begin
      if TobParSuspect.GetValue('RSS_SUFFIXE') = 'MAJ' then
        bMajCodeTiers   := True
      else
        bMajCodeTiers   := False;
        
      RSS_FICHIER.Text := TobParSuspect.GetValue ('RSS_FICHIER');
      LRSS_PARSUSPECT.Caption := TobParSuspect.GetValue ('RSS_LIBELLE');
      if (TobParSuspect.GetValue ('RSS_TYPEENREG') = 'X') then
      LongFixe := True else LongFixe := False;
      RempliCombo;
    end;
    if (getPage = Apercu) then InitialiseApercu;
  end;
end;

function TFRecupSuspect.NextPage: TTabSheet;
begin
  if not passage then begin passage := true; result := Fichier; exit; end;
  if (getPage = Fichier) then
  begin
    result := Fichier;
    if  (TobParSuspectLig.detail.count>0) and FileExists(RSS_FICHIER.Text) then
    begin
      if VerifAccesFichier = True then
      begin
        InitialiseApercu;
        result := Apercu;
      end else result := Apercu
    end
  end else
  if P.ActivePage.PageIndex<P.PageCount-1 then result:=P.Pages[P.ActivePage.PageIndex+1]
  else result:=nil ;
  if (getPage = Criteres) and (not bMajCodeTiers) then
    ListeRecap
  else
    ListRecapMajCodeTiers;
  if (getPage = Recapitulatif)  then bfin.enabled:=true;
end;

procedure TFRecupSuspect.bPrecedentClick(Sender: TObject);
begin
  inherited;
 bfin.enabled := not bSuivant.Enabled;
end;

procedure TFRecupSuspect.bFinClick(Sender: TObject);
begin
  inherited;
if getPage <> Recapitulatif then exit;
if ControleSaisie then
    begin
    if fileexists (RSS_FICHIER.Text) then
        begin
        if VerifAccesFichier = True then
            begin
            bfin.enabled := false;
            if not bMajCodeTiers then
              RecupereSuspect
            else
              MajTiersParCode;
            if BStop = False then
                begin
                Msg.Execute(9, Caption, '');
    //            Close;
                end;
            end;
        end else Msg.Execute (7, Caption, '');
    end else ModalResult := 0;
end;

procedure TFRecupSuspect.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
if (RSS_PARSUSPECT.Text <> '') then
begin
TobParSuspect.PutValue ('RSS_PREFIXE', OrigineFichier.value);
TobParSuspect.PutValue ('RSS_FICHIER', RSS_FICHIER.Text);
TobParSuspect.SetAllModifie (True);
TobParSuspect.UpdateDB;
end;
DepileTOBLigne;
TobParSuspectLig.Free;
TobParSuspect.Free;
TobSuspect.Free;
TobSuspectCompl.Free;
TobTablette.Free;
TobCorresp.Free;
end;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

procedure TFRecupSuspect.InitialiseApercu;
var LesColonnes : string;
    Fichier : textfile;

begin
  GChamp.Rowcount := TobParSuspectLig.detail.count+1;
  GChamp.ColWidths [SAF_Libelle] := 200;
  GChamp.ColWidths [SAF_Offset] := 90;
  GChamp.ColAligns[SAF_offset]:=taRightJustify;
  if (LongFixe) then
  begin
  GChamp.ColWidths [SAF_Longueur] := 90;
  GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Début');
  end else
  begin
  GChamp.ColWidths[SAF_Longueur] := 0;
  GChamp.Cells [SAF_Offset,0] := TraduireMemoire('Position');
  end;
  GChamp.ColAligns[SAF_Longueur]:=taRightJustify;
  GChamp.ColWidths [SAF_Champ] := 200;
  GChamp.ColWidths [SAF_Champ2] := 200;
  GChamp.ColWidths [SAF_Champ3] := 200;
  AssignFile (Fichier, RSS_FICHIER.Text);
  Reset (Fichier);
  readln (Fichier,stEnreg);     // lecture des données
  if CBPremiereLigne.checked then
    readln (Fichier,stEnreg);
  TobParSuspectLig.Detail[0].AddChampSup('LIGNE1', True);
  TraiteEnregistrementApercu ('LIGNE1') ;
  readln (Fichier,stEnreg);
  TobParSuspectLig.Detail[0].AddChampSup('LIGNE2', True);
  TraiteEnregistrementApercu ('LIGNE2') ;
  readln (Fichier,stEnreg);
  TobParSuspectLig.Detail[0].AddChampSup('LIGNE3', True);
  TraiteEnregistrementApercu ('LIGNE3') ;
  LesColonnes :='RRL_LIBELLE;RRL_OFFSET;RRL_LONGUEUR;LIGNE1;LIGNE2;LIGNE3';

  TobParSuspectLig.PutGridDetail(GChamp,False,False,LesColonnes,true);
  HMTrad.ResizeGridColumns (GChamp) ;
  CloseFile(Fichier);
end;

procedure TFRecupSuspect.TraiteEnregistrementApercu ( Ligne : string);
var Index : integer;
    ValChamp, Champ : string;
    LeResultat : String;
    LaFormule : String;

begin
TobSuspect.InitValeurs;
TobSuspectCompl.InitValeurs;
for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
  begin
    Champ := TobParSuspectLig.Detail[Index].GetValue ('RRL_CHAMP');
    RecupereValChamp (TobParSuspectLig.Detail[Index], Index, Valchamp);
    LaFormule := TobParSuspectLig.Detail[Index].GetValue('RRL_FORMULE');
    if Trim(LaFormule) <> '' then
    begin
      LeResultat := TraduireFormule(LaFormule, Champ);
      ValChamp := LeResultat;
    end;
    TobParSuspectLig.Detail[Index].PutValue(Ligne, Valchamp);
  end;
end;

procedure TFRecupSuspect.ActiveChamp (HCrit : THCritMaskEdit);
begin
HCrit.Enabled := True;
HCrit.Color := clWindow;
end;

Procedure TFRecupSuspect.DepileTOBLigne ;
var Index : integer;
BEGIN
for Index := TobParSuspectLig.Detail.Count - 1 Downto 0 do
    BEGIN
    TobParSuspectLig.Detail[Index].Free ;
    end;
end;

procedure TFRecupSuspect.DesactiveChamp (HCrit : THCritMaskEdit);
begin
HCrit.Enabled := False;
HCrit.Color := clActiveBorder;
HCrit.Text := '';
end;

Procedure TFRecupSuspect.InitialiseFiche;
begin
  if (LoadLesTobParSuspect (RSS_PARSUSPECT.Text, TobParSuspect, TobParSuspectLig)) then
  begin
    if TobParSuspect.GetValue('RSS_SUFFIXE') = 'MAJ' then
      bMajCodeTiers     := True;
    RSS_FICHIER.Text := TobParSuspect.GetValue ('RSS_FICHIER');
    LRSS_PARSUSPECT.Caption := TobParSuspect.GetValue ('RSS_LIBELLE');
    OrigineFichier.value := TobParSuspect.GetValue ('RSS_PREFIXE');
    if (TobParSuspect.GetValue ('RSS_TYPEENREG') = 'X') then
    LongFixe := True else LongFixe := False;
    if (TobParSuspect.GetValue ('RSS_SUFFIXE') = 'PAR') then Particulier:=true
    else Particulier:=false;
    RempliCombo;
  end;
  PTITRE.Caption        := 'Import d''un fichier Suspects';
  CBCreationSuspect.Checked := False;
  CBPremiereLigne.Checked := True;
  HChamp1.Text := stAucun;
  HChamp2.Enabled := False;
  HChamp3.Enabled := False;
  HChamp2.Text := stAucun;
  HChamp3.Text := stAucun;

  DesactiveChamp (HLongueur1);
  DesactiveChamp (HLongueur2);
  DesactiveChamp (HLongueur3);

  DesactiveChamp (HOffset1);
  DesactiveChamp (HOffset2);
  DesactiveChamp (HOffset3);

  DesactiveChamp (HBorneInf1);
  DesactiveChamp (HBorneInf2);
  DesactiveChamp (HBorneInf3);

  DesactiveChamp (HBorneSup1);
  DesactiveChamp (HBorneSup2);
  DesactiveChamp (HBorneSup3);
  CompteurCode :='';

  //FQ10505
  //on enleve les tiers de type Autre si pas ctxScot (GI)
  if (ctxScot in V_PGI.PGIContexte) then
    CbAutre.Visible := True
  else
    CbAutre.Visible := False;
  // Concept de modification des Tiers
  if (not ExJaiLeDroitConcept(TConcept(gcImportModifTiers), False)) then
  begin
    CBMajTiers.Visible := False;
    GBTiers.Visible := False;
    CBContact.Visible := False;
  end;

  if bMajCodeTiers then
  begin
    CBCreationSuspect.Visible   := False;
    CBControleProspect.Visible  := False;
    CBMajTiers.Visible          := False;
    CBContact.Visible           := False;
    GBTiers.Visible             := False;
    CBtest.Visible              := False;
    PTITRE.Caption              := 'Mise à jour des fiches tiers';

  end;

end;

procedure TFRecupSuspect.RempliCombo;
var Index, IndexItem : integer;                                               
    LibelleChamp : string;
begin
HChamp1.values.clear; HChamp2.values.clear;  HChamp3.values.clear;
HChamp1.items.clear;  HChamp2.items.clear;   HChamp3.items.clear;
for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
    begin
    LibelleChamp := TobParSuspectLig.Detail[Index].GetValue ('RRL_LIBELLE');
    IndexItem := HChamp1.Items.Add (LibelleChamp);
    HChamp1.Items.objects[IndexItem] := TobParSuspectLig.Detail [Index];
    IndexItem := HChamp2.Items.Add (LibelleChamp);
    HChamp2.Items.Objects[IndexItem] := TobParSuspectLig.Detail [Index];
    IndexItem := HChamp3.Items.Add (LibelleChamp);
    HChamp3.Items.objects[IndexItem] := TobParSuspectLig.Detail [Index];
    end;

HChamp1.Items.add (stAutre);
HChamp2.Items.add (stAutre);
HChamp3.Items.add (stAutre);
HChamp1.Items.add (stAucun);
HChamp2.Items.add (stAucun);
HChamp3.Items.add (stAucun);
end;


procedure TFRecupSuspect.HChamp1Change(Sender: TObject);
begin
  inherited;
if HChamp1.Text = stAucun then
    begin
    DesactiveChamp (HBorneInf1);
    DesactiveChamp (HBorneSup1);
    DesactiveChamp (HLongueur1);
    DesactiveChamp (HOffset1);
    HChamp2.Enabled := False;
    HChamp2.Text := stAucun;
    HChamp2Change (HChamp1);
    end else
    begin
    HChamp2.Enabled := True;
    if HChamp1.Text = stAutre then
        begin
        if LongFixe then ActiveChamp (HLongueur1);
        ActiveChamp (HOffset1);
        ActiveChamp (HBorneInf1);
        ActiveChamp (HBorneSup1);
        end else
        begin
        ActiveChamp (HBorneInf1);
        ActiveChamp (HBorneSup1);
        DesactiveChamp (HLongueur1);
        DesactiveChamp (HOffset1);
        end;
    end;
end;

procedure TFRecupSuspect.HChamp2Change(Sender: TObject);
begin
  inherited;
if HChamp2.Text = stAucun then
    begin
    DesactiveChamp (HBorneInf2);
    DesactiveChamp (HBorneSup2);
    DesactiveChamp (HLongueur2);
    DesactiveChamp (HOffset2);
    HChamp3.Enabled := False;
    HChamp3.Text := stAucun;
    HChamp3Change (HChamp2);
    end else
    begin
    HChamp3.Enabled := True;
    if HChamp2.Text = stAutre then
        begin
        if LongFixe then ActiveChamp (HLongueur2);
        ActiveChamp (HOffset2);
        ActiveChamp (HBorneInf2);
        ActiveChamp (HBorneSup2);
        end else
        begin
        ActiveChamp (HBorneInf2);
        ActiveChamp (HBorneSup2);
        DesactiveChamp (HLongueur2);
        DesactiveChamp (HOffset2);
        end;
    end;
end;

procedure TFRecupSuspect.HChamp3Change(Sender: TObject);
begin
  inherited;
if HChamp3.Text = stAucun then
    begin
    DesactiveChamp (HBorneInf3);
    DesactiveChamp (HBorneSup3);
    DesactiveChamp (HLongueur3);
    DesactiveChamp (HOffset3);
    end else
    begin
    if HChamp3.Text = stAutre then
        begin
        if LongFixe then ActiveChamp (HLongueur3);
        ActiveChamp (HOffset3);
        ActiveChamp (HBorneInf3);
        ActiveChamp (HBorneSup3);
        end else
        begin
        ActiveChamp (HBorneInf3);
        ActiveChamp (HBorneSup3);
        DesactiveChamp (HLongueur3);
        DesactiveChamp (HOffset3);        
        end;
    end;
end;

{==============================================================================================}
{=============================== Controles ====================================================}
{==============================================================================================}

function TFRecupSuspect.ControleChampOk (HChamp : THValcomboBox; ValChamp : string;
                                           stLongueur, stOffset, stBorneInf, stBorneSup : string) : boolean;
var IndexItem : integer;
begin
  Result := True;
  if HChamp.Text <> stAutre then
      begin
        IndexItem := HChamp.Items.IndexOf (HChamp.Text);
        RecupereValChamp (TOB(HChamp.Items.Objects[IndexItem]), IndexItem, ValChamp);
      end else
      begin
        if (LongFixe) then ValChamp := Copy (stEnreg, StrToInt (stOffset), StrToInt (stLongueur))
        else RecupValChampEnregVar (ValChamp, StrToInt(stOffset)-1);
      end;
if (Copy(ValChamp, 0, Length(stBorneInf)) < stBorneInf) or
   (Copy(ValChamp, 0, Length(stBorneSup)) > stBorneSup) then Result := False;
end;

function TFRecupSuspect.ControleEnreg : boolean;
var ValChamp : string;
begin
Result := True;
if HChamp1.Text <> stAucun then
    begin
    if not ControleChampOk (HChamp1, ValChamp, HLongueur1.Text, HOffset1.Text,
                            HBorneInf1.Text, HBorneSup1.Text) then Result := False
    else if HChamp2.Text <> stAucun then
             begin
             if not ControleChampOk (HChamp2, ValChamp, HLongueur2.Text, HOffset2.Text,
                                     HBorneInf2.Text, HBorneSup2.Text) then Result := False
             else if HChamp3.Text <> stAucun then
                      begin
                      if not ControleChampOk (HChamp3, ValChamp, HLongueur3.Text, HOffset3.Text,
                                              HBorneInf3.Text, HBorneSup3.Text) then Result := False;
                      end;
             end;
    end;
end;

function TFRecupSuspect.ControleSaisie : boolean;
begin
Result := True;
if not ControleSelection (HChamp1.Text, HBorneSup1.Text, HBorneInf1.Text, HLongueur1.Text,
                          HOffset1.Text, 1, 8) then Result := False
    else if not ControleSelection (HChamp2.Text, HBorneSup2.Text, HBorneInf2.Text, HLongueur2.Text,
                                   HOffset2.Text, 3, 4) then Result := False
         else if not ControleSelection (HChamp3.Text, HBorneSup3.Text, HBorneInf3.Text,
                                        HLongueur3.Text, HOffset3.Text, 5, 6) then Result := False;
end;

function TFRecupSuspect.ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur,
                                             stOffset : string; Message1, Message2 : integer) : boolean;
begin
Result := True;
if stChamp <> stAucun then
    begin
    if stBorneSup < stBorneInf then
        begin
        Msg.Execute (Message1, Caption, '');
        Result := False;
        end else
        begin
        if stChamp = stAutre then
            begin
              if ( not LongFixe and (stOffset = '') ) or  ( LongFixe and ((stLongueur = '') or (stOffset = '')) ) then
                begin
                Msg.Execute (Message2, Caption, '');
                Result := False;
                end;
            end;
        end;
    end;
end;

{==============================================================================================}
{=============================== Traitements ==================================================}
{==============================================================================================}

procedure TFRecupSuspect.CompleteChampSuspect;
var
  CodeSuspect : String;
//  titre : string;
begin

  CodeSuspect := AffectCodeSuspect;
  TobSuspect.PutValue ('RSU_SUSPECT',CodeSuspect);
  Tobsuspectcompl.PutValue('RSC_SUSPECT',CodeSuspect);
 {
  TobSuspect.ChargeCle1;
  if TobSuspect.existDB then
  begin
    titre :=  TraduireMemoire('Code suspect existant : ') + TobSuspect.getValue ('RSU_SUSPECT');
    ListRecap.Items.Add (titre);
    BSTop := (Msg.Execute(10, titre, '') <> mrYes);
    if not BStop then CompleteChampSuspect ;
  end;
 }
  TobSuspect.PutValue ('RSU_FERME', '-');
  TobSuspect.PutValue ('RSU_DATECREATION', V_PGI.DateEntree);
  TobSuspect.PutValue ('RSU_DATEMODIF', V_PGI.DateEntree);
  TobSuspect.PutValue ('RSU_ORIGINETIERS', OrigineFichier.value );
  if (trim(TobSuspect.GetValue ('RSU_PAYS'))='') then TobSuspect.PutValue ('RSU_PAYS', 'FRA');
  if (trim(TobSuspect.GetValue ('RSU_LANGUE'))='') and (trim(TobSuspect.GetValue ('RSU_PAYS'))='FRA')
  then TobSuspect.PutValue ('RSU_LANGUE', 'FRA');
  if (trim(TobSuspect.GetValue ('RSU_NATIONALITE'))='') and (trim(TobSuspect.GetValue ('RSU_PAYS'))='FRA')
  then TobSuspect.PutValue ('RSU_NATIONALITE', 'FRA');
  if Particulier then TobSuspect.PutValue ('RSU_PARTICULIER', 'X')
  else TobSuspect.PutValue ('RSU_PARTICULIER', '-');
if VH_GC.GCIfDefCEGID then
  if (OrigineFichier.value = '001') then
  begin
    if (trim(TobSuspect.GetValue ('RSU_JURIDIQUE'))<>'') then CorrespondanceJur;
    if (trim(TobSuspect.GetValue ('RSU_CONTACTFONCTION'))<>'') then CorrespondanceFonc;
    CorrespondanceEff;
    CorrespondanceCA;
    CorrespondanceDateCreation ;
  end;
end;

function TFRecupSuspect.AffectCodeSuspect :string;
begin
  if (trim(compteurCode) = '') then
    compteurCode        := GetParamsocSecur('SO_RTCOMPTEURSUSP', 0, True);
  //appel à une fonction de recherche du dernier code suspect
    compteurCode          := RTAttribNewCodeSuspect (compteurCode);
    Result                := compteurCode;
end;

procedure TFRecupSuspect.ListeRecap;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (TraduireMemoire('Récapitulatif: ') + TraduireMemoire(PTITRE.Caption));
ListRecap.Items.Add (TraduireMemoire('Fichier: ') + RSS_FICHIER.Text);

if CBCreationSuspect.checked then
    BEGIN
    ListRecap.Items.Add (ExtractLibelle (CBCreationSuspect.Caption) + HRecap.Mess[0]);
    end else
    BEGIN
    ListRecap.Items.Add (ExtractLibelle (CBCreationSuspect.Caption) + HRecap.Mess[1]);
    end;

if CBControleProspect.checked then
    BEGIN
    ListRecap.Items.Add (ExtractLibelle (CBControleProspect.Caption) + HRecap.Mess[0]);
    end else
    BEGIN
    ListRecap.Items.Add (ExtractLibelle (CBControleProspect.Caption) + HRecap.Mess[1]);
    end;

if Particulier then
    ListRecap.Items.Add (TraduireMemoire('Création des fiches en Suspect Particulier'))
    else ListRecap.Items.Add (TraduireMemoire('Création des fiches en Suspect Entreprise'));

if HChamp1.Text <> stAucun then
    begin
    ListRecap.Items.Add (TraduireMemoire('Critères de sélection des enregistrements du fichier'));
    RecapChamp (HChamp1.Text, HLongueur1.Text, HOffset1.Text, HBorneInf1.Text,
                HBorneSup1.Text);
    if HChamp2.Text <> stAucun then
        begin
        ListRecap.Items.Add ('');
        RecapChamp (HChamp2.Text, HLongueur2.Text, HOffset2.Text, HBorneInf2.Text,
                    HBorneSup2.Text);
        if HChamp3.Text <> stAucun then
            begin
            ListRecap.Items.Add ('');
            RecapChamp (HChamp3.Text, HLongueur3.Text, HOffset3.Text, HBorneInf3.Text,
                        HBorneSup3.Text);
            end;
        end;
    end else
    begin
    ListRecap.Items.Add (TraduireMemoire('Pas de critères de sélection des enregistrements du fichier'));
    end;
//ajout au récap des nouvelles fonctionnalités
if CBMajTiers.Checked then
begin
   ListRecap.Items.Add(ExtractLibelle(CBMajTiers.Caption) + HRecap.Mess[0]);
  if CbProspect.Checked then
    ListRecap.Items.Add(ExtractLibelle(CbProspect.Caption) + HRecap.Mess[0]);
  if CbClient.Checked then
    ListRecap.Items.Add(ExtractLibelle(CbClient.Caption) + HRecap.Mess[0]);
  if CbAutre.Checked then
    ListRecap.Items.Add(ExtractLibelle(CbAutre.Caption) + HRecap.Mess[0]);
end
else
   ListRecap.Items.Add(ExtractLibelle(CBMajTiers.Caption) + HRecap.Mess[1]);

ListRecap.Items.Add ('');

// mise à jour des contacts
if CBContact.Checked then
  ListRecap.Items.Add(ExtractLibelle(CBContact.Caption) + HRecap.Mess[0])
else
  ListRecap.Items.Add(ExtractLibelle(CBContact.Caption) + HRecap.Mess[1]);

end;

procedure TFRecupSuspect.RecapChamp (stChamp, stLongueur, stOffset, stBorneInf,
                                       stBorneSup : string);
var stChaine : string;
begin
if stChamp <> stAutre then
    begin
    ListRecap.Items.Add (TraduireMemoire('Sélection sur le champ : ') + stChamp);
    end else
    begin
    ListRecap.Items.Add (TraduireMemoire('Sélection à la position : ') + stOffset);
    if (stLongueur<>'') then
      begin
      if (Valeuri(stLongueur) > 1) then stChaine := Format(TraduireMemoire('    sur une longueur de : %s caractères'),[stlongueur])
      else stChaine := Format(TraduireMemoire('    sur une longueur de : %s caractère'),[stlongueur]);
      ListRecap.Items.Add (stChaine);
      end;
    end;
ListRecap.Items.Add (TraduireMemoire('    Borne inférieure : ') + stBorneInf);
ListRecap.Items.Add (TraduireMemoire('    Borne supérieure : ') + stBorneSup);
end;

procedure TFRecupSuspect.RecupereSuspect;
var
  NligneFichier         : longint;
  Fichier               : textfile;
  ComplSql              : string;
  ioerr                 : TIOErr;
  Save_Cursor           : TCursor;
  Nb_MajFiche           : longint;
//  GereEnseigne          : Boolean;
  BRollInsert           : boolean;
  BrollDoublon          : Boolean;
  BRollOrigine          : Boolean;
  Nb_Block              : integer;
  PremierCode           : integer;
  DernierSuspect        : String;
  ValMajTiers           : integer;
  StrSql                : String;

begin
  Nb_MajFiche           := 0;
  PremierCode           := 0;
  BoAjoutCorr           := False;
  Save_Cursor           := Screen.Cursor;
  Screen.Cursor         := crHourglass;
  BStop                 := False;
  ioerr                 := oeOk;
  Nb_Ecrit              := 0;
  Nb_Block              := 0;
  Nb_ErreurSiret        := 0;
  BRollInsert           := False;
  BrollDoublon          := False;
  BRollOrigine          := False;
  NligneFichier         := InitProgression ;
  InitMoveProgressForm (nil,TraduireMemoire('Traitement en cours...'),'',NligneFichier,TRUE,TRUE) ;
  ListRecap.Items.Add (TraduireMemoire('Nombre de lignes du fichier : ') + IntToStr(NligneFichier));
  AssignFile (Fichier, RSS_FICHIER.Text); // ouvre le pointeur sur le fichier
  Reset (Fichier);
//  GereEnseigne          := GetParamSocSecur('SO_GCENSEIGNETAB', False, True);

  try   //finally
  // on suspprime les supects si code origne identique
    if CBCreationSuspect.checked then
    begin
      Bstop             := not(MoveCurProgressForm (TraduireMemoire('Efface la table des suspects')));
      BeginTrans;
      ExecuteSQL('Delete from suspectscompl WHERE RSC_SUSPECT in (select RSU_SUSPECT from suspects WHERE RSU_ORIGINETIERS="' + OrigineFichier.value + '"  )');
      ExecuteSQL('Delete from SUSPECTS where RSU_ORIGINETIERS ="' + OrigineFichier.value + '"  ');
      try
        CommitTrans;
      except
        on e:exception do
        begin
          RollBack;
          BRollOrigine  := True;
        end;
      end;
    end;

    Bstop               := not(MoveCurProgressForm (TraduireMemoire('Récupération Fichier suspects')));
    if CBPremiereLigne.checked then
      readln (Fichier,stEnreg);
    ChargeTablette;
    ChargeTableCorresp;
    ChercheProchainCode;
    Nb_Block            := 0;

    while not EOF(Fichier) and not Bstop do
    begin
      Application.ProcessMessages;
      if Not MoveCurProgressForm(TraduireMemoire('Récupération Fichier suspects')) then
        break ;
      readln (Fichier,stEnreg);     // lecture des données
      if ControleEnreg then
        try
          if (not CBtest.Checked) and ((Nb_Ecrit mod KBlock) = 0) then
            BEGINTRANS;
          TraiteEnregistrement (False);
          if (not CBtest.Checked) and (Nb_Ecrit = 0) then
            PremierCode := TobSuspect.GetValue('RSU_SUSPECT');

          if not BStop then   //si on arrete pas la procèdure
          begin
            if not CbTest.Checked then    //Mode Test pour maj table de correspondance
            begin
              EnregLesSuspects;
              if (CBMajTiers.Checked) and (TobSuspect.GetValue('RSU_SIRET')<>'') then //si maj Tiers et Code siret existant
              begin
                ComplSql := '';
                ValMajTiers := 0;
                if CbProspect.Checked then
                  Inc(ValMajTiers, 1);
                if CbClient.Checked then
                  Inc(ValMajTiers, 2);
                if CbAutre.Checked then
                  Inc(ValMajTiers, 4);
                Case ValMajTiers of
                  1 : ComplSql := ' AND T_NATUREAUXI="PRO"';
                  2 : ComplSql := ' AND T_NATUREAUXI="CLI"';
                  3 : ComplSql := ' AND (T_NATUREAUXI="PRO" OR T_NATUREAUXI="CLI")';
                  4 : ComplSql := ' AND T_NATUREAUXI="NCP"';
                  5 : ComplSql := ' AND (T_NATUREAUXI="PRO" OR T_NATUREAUXI="NCP")';
                  6 : ComplSql := ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="NCP")';
                  7 : ComplSql := ' AND (T_NATUREAUXI="PRO" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="NCP")';
                end;
                StrSql := 'SELECT 1 FROM TIERS WHERE T_SIRET="'+TobSuspect.getvalue('RSU_SIRET')+'"' +
                          ' AND T_FERME = "-"' +  //rajout Email M Scheidegger
                          ComplSql;
                if (valmajtiers <> 0) and (ExisteSQL(StrSql)) then
                begin
                // fonction de maj tiers
                // ajout de l'option maj contact
                  if SuspectVersProspect(TobSuspect.getvalue('RSU_SUSPECT'), False, False, CBContact.Checked) <> '' then
                    inc (Nb_MajFiche);
                end;
              end;
            end;
            inc(Nb_Ecrit);

          end;
        if ((Nb_Ecrit mod KBlock) = 0) and (not CBtest.Checked) then
        begin
          SetParamSoc('SO_RTCOMPTEURSUSP', compteurCode);
          COMMITTRANS;
          inc (Nb_Block, KBlock);
          DernierSuspect  := TobSuspect.GetValue('RSU_LIBELLE');
//          ListRecap.Items.Add('Compteur : ' + IntToStr(Nb_Ecrit));
        end;
        except
          on e:exception do
          begin
            ROLLBACK;
            BRollInsert := True;
            Break;
          end;
        end;
      end;
      CloseFile (Fichier);

      Try
      //cas où on n'est pas au bout du block et pas mode test
      if ((Nb_Ecrit mod KBlock) <> 0) and (Not BRollInsert) and (not CBtest.Checked) then
        begin
          SetParamSoc('SO_RTCOMPTEURSUSP', compteurCode) ;
          CommitTrans;
          Nb_Block      := Nb_Ecrit;
        end;


        TobCorresp.InsertOrUpdateDB(False);   // maj table correspondance
      Except
          on e:exception do
          begin
            RollBack ;
            BRollInsert := True;
          end;
      end;





      if not CBtest.Checked then
      begin
        if not Bstop and CBControleProspect.checked then
        begin
          MoveCurProgressForm (TraduireMemoire('Efface les suspects existants'));
          try
          BeginTrans;
          ExecuteSQL('DELETE FROM SUSPECTS WHERE RSU_SIRET <> "" AND ' +
                     'RSU_ORIGINETIERS ="' + OrigineFichier.value + '" AND '+
                     'EXISTS ( SELECT T_AUXILIAIRE FROM TIERS WHERE T_SIRET=SUSPECTS.RSU_SIRET ) ');
          ExecuteSQL('DELETE FROM SUSPECTSCOMPL WHERE '+
          'NOT EXISTS ( SELECT RSU_SUSPECT FROM SUSPECTS WHERE RSU_SUSPECT=SUSPECTSCOMPL.RSC_SUSPECT ) ');
          CommitTrans;
          except
            on e:exception do
            begin
              RollBack ;
              BrollDoublon  := True;
            end;
          end;
        end;
      end;

      // maj date import dans tous les cas
      TobParSuspect.PutValue ('RSS_UTILISATEUR', V_PGI.User) ;
      TobParSuspect.UpdateDateModif;

    Finally
      //Gestion du journal
      if not CBtest.Checked then
      begin
        ListRecap.Items.Add('Premier code suspect : ' + IntToStr(PremierCode));
        ListRecap.Items.Add(TraduireMemoire('Dernier code suspect : ')+compteurCode);
        ListRecap.Items.Add('Nombre de lignes suspects lues : ' + IntToStr(Nb_Ecrit));
        ListRecap.Items.Add(TraduireMemoire('Nombre de lignes suspects écrites : ') + IntToStr(Nb_Block));
        if CBMajTiers.Checked then
          ListRecap.Items.Add('Nombre de fiches tiers mises à jour : ' + IntToStr(Nb_MajFiche));
      end;

      //Gestion des erreurs pour le journal
      if (Nb_ErreurSiret > 0) then
      begin
        ListRecap.Items.Add(IntToStr(Nb_ErreurSiret)+TraduireMemoire(' codes Siret ne sont pas corrects et donc pas renseignés '));
      end;

      if BRollInsert then
      begin
        ListRecap.Items.Add('==================================');
        ListRecap.Items.Add('!!Erreur sur l''import');
        if V_PGI.IOError = oeReseau then
        begin
          ListRecap.Items.Add('Problème réseau');
          ListRecap.Items.Add('Dernière ligne enregistrée : ' + IntToStr(Nb_Block));
          ListRecap.Items.Add('Suspect : ' + DernierSuspect);
        end
        else
        begin
          ListRecap.Items.Add('Arrêt sur la ligne du fichier import : ' + IntToStr(Nb_Ecrit));
          ListRecap.Items.Add('Suspect : ' + TobSuspect.GetValue('RSU_LIBELLE'));
        end;
        ListRecap.Items.Add('==================================');
      end;

      if BRollOrigine then
        ListRecap.Items.Add('Une erreur est apparue à la suppression des suspects ayant le même code origine');

      if BrollDoublon then
        ListRecap.Items.Add('Une erreur est apparue à la suppression des suspects ayant le même code siret qu''un tiers');


      FiniMoveProgressForm ;
      Screen.Cursor := Save_Cursor;
      RenseigneJournalEvenement (ioerr);

    end;
    // si on ajoute une correspondance
    If (Cbtest.Checked) and (BoAjoutCorr) then
    begin
      ParamTable('RTCORRESPIMPORT', taModif, -1, nil);
    end;
end;

procedure TFRecupSuspect.TraiteEnregistrement (Apercu : boolean);
var Index : integer;
    ValChamp, Champ : string;
    LeResultat : String;
    LaFormule : String;

begin
TobSuspect.InitValeurs;
TobSuspectCompl.InitValeurs;
for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
  begin
    try
        Champ := TobParSuspectLig.Detail[Index].GetValue ('RRL_CHAMP');
        RecupereValChamp (TobParSuspectLig.Detail[Index], Index, Valchamp);
        LaFormule := TobParSuspectLig.Detail[Index].GetValue('RRL_FORMULE');
        if Trim(LaFormule) <> '' then
        begin
          LeResultat := TraduireFormule(LaFormule, Champ);
          ValChamp := LeResultat;
        end;
        TraiteChamp (Index, ValChamp, Champ);
    except
      on e: exception do
         ListRecap.Items.Add ('erreur champ :' +IntToStr(index)+' val='+ValChamp+' ch='+Champ );

    end;
  end;
  if not CBtest.Checked then
    CompleteChampSuspect;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 04/01/2007
Modifié le ... :   /  /    
Description .. : enregistement de Suspect et SuspectCompl
Mots clefs ... : 
*****************************************************************}
procedure TFRecupSuspect.EnregLesSuspects;
begin
  TobSuspect.SetAllModifie(True);
  TobSuspect.ChargeCle1;
  TobSuspectCompl.SetAllModifie(True);
  TobSuspectCompl.ChargeCle1;

  TobSuspect.InsertDB(nil, False);
  TobSuspectCompl.InsertDB(nil, False);
{  except
    on e:exception do
    begin
      ListRecap.Items.Add('Erreur sur l''enregistrement ' + IntToStr(Nb_Ecrit) + ' : ' + TobSuspect.GetValue('RSU_LIBELLE'));
      ListRecap.Items.Add ('erreur : '+e.message);
    end;
  end;
 }
end;

(*
procedure TFRecupSuspect.ImportEnregSuspect;
var
//    vt_trav1 : variant;
//    st_trav1, CleTable, OrdreSQL : string;
    CleTable : string;
//    i_ind1 : integer;
begin
  SQL_Champs := '';
  SQL_Datas  := '';
  SQL_Where  := '';
  SQL_Sep    := '';
  SQL_Insert := 'Insert into SUSPECTS';
  SQL_Update := 'Update SUSPECTS set ';
  TobSuspect.SetAllModifie(True);
  TobSuspect.ChargeCle1;
  CleTable := TobSuspect.Cle1;
{  if (CBMaj.checked and TobSuspect.existDB) then
  begin
    for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
    begin
      st_trav1 := TobParSuspectLig.Detail[Index].GetValue ('RRL_CHAMP');
      if ((st_trav1<>'RSU_SUSPECT') and TobSuspect.FieldExists (st_trav1)) then
      begin
      i_ind1 := TobSuspect.GetNumChamp(st_trav1);
      vt_trav1 := TobSuspect.GetValeur(i_ind1);
      Ecrire_Datas(st_trav1, TobSuspect.NumTable, i_ind1,
                   vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep);
      end;
    end;
    SQL_Sep := '';
    OrdreSQL := SQL_Update;
    while SQL_Champs <> '' do
    begin
      OrdreSQL := OrdreSQL + SQL_Sep + Trim(ReadTokenStV(SQL_Champs)) +
                    '=' + Trim(ReadTokenStV(SQL_Datas));
      SQL_Sep := ',';
    end;
    OrdreSQL := OrdreSQL + ' where ' + CleTable;
  end else   }

// Insert into ... avec bloc-note plantait avec DB2
{  for i_ind1 := 1 to TobSuspect.NbChamps do
  begin
    st_trav1 := TobSuspect.GetNomChamp(i_ind1);
    vt_trav1 := TobSuspect.GetValeur(i_ind1);
    Ecrire_Datas(st_trav1, TobSuspect.NumTable, i_ind1,
                 vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep);
  end;
  OrdreSQL := SQL_Insert + ' (' + SQL_Champs + ') values (' + SQL_Datas + ')';

  ExecuteSQL(OrdreSQL);   }
  TobSuspect.InsertDB(Nil);

  if (Nb_Ecrit mod 5) = 0 then
  begin
    Try
    if (Nb_Ecrit > 0) then
    begin
        CommitTrans;
        ListRecap.Items.Add (TraduireMemoire('Nombre de lignes suspects écrites: ') + IntToStr(Nb_Ecrit));
        SetParamSoc('SO_RTCOMPTEURSUSP', compteurCode) ;
    end else
        ListRecap.Items.Add (TraduireMemoire('Premier code suspect : ')+compteurCode);
    Except
          on e:exception do
          begin
            ListRecap.Items.Add ('erreur : '+e.message);
            RollBack ;
          end;
    end;
    BeginTrans;
  end;
end;


procedure TFRecupSuspect.ImportEnregSuspectCompl;
var
//    vt_trav1 : variant;
//    st_trav1 : string;
    CleTable : string;
//    OrdreSQL : string;
//    i_ind1 : integer;
begin
  SQL_Champs := '';
  SQL_Datas  := '';
  SQL_Where  := '';
  SQL_Sep    := '';
  SQL_Insert := 'Insert into SUSPECTSCOMPL';
  SQL_Update := 'Update SUSPECTSCOMPL set ';
  TobSuspectCompl.SetAllModifie(True);
  TobSuspectCompl.ChargeCle1;
  CleTable := TobSuspectCompl.Cle1;
{  if (CBMaj.checked and TobSuspectCompl.existDB) then
  begin
    for Index := 0 to TobParSuspectLig.Detail.Count - 1 do
    begin
      st_trav1 := TobParSuspectLig.Detail[Index].GetValue ('RRL_CHAMP');
      if ((st_trav1<>'RSC_SUSPECT') and TobSuspectCompl.FieldExists (st_trav1)) then
      begin
      i_ind1 := TobSuspectCompl.GetNumChamp(st_trav1);
      vt_trav1 := TobSuspectCompl.GetValeur(i_ind1);
      Ecrire_Datas(st_trav1, TobSuspectCompl.NumTable, i_ind1,
                   vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep);
      end;
    end;
    SQL_Sep := '';
    OrdreSQL := SQL_Update;
    while SQL_Champs <> '' do
    begin
      OrdreSQL := OrdreSQL + SQL_Sep + Trim(ReadTokenStV(SQL_Champs)) +
                    '=' + Trim(ReadTokenStV(SQL_Datas));
      SQL_Sep := ',';
    end;
    OrdreSQL := OrdreSQL + ' where ' + CleTable;
  end else     }


 {
  for i_ind1 := 1 to TobSuspectCompl.NbChamps do
  begin
    st_trav1 := TobSuspectCompl.GetNomChamp(i_ind1);
    vt_trav1 := TobSuspectCompl.GetValeur(i_ind1);
    Ecrire_Datas(st_trav1, TobSuspectCompl.NumTable, i_ind1,
                 vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep);
  end;
  OrdreSQL := SQL_Insert + ' (' + SQL_Champs + ') values (' + SQL_Datas + ')';

  ExecuteSQL(OrdreSQL);
}
//ciblage : enreg par la tob et non comme si dessus avec un executeSQL
  TobSuspectCompl.InsertDB(Nil);

  if (Nb_Ecrit mod 5) = 0 then
  begin
    Try
        CommitTrans;
    Except
          on e:exception do
          begin
            ListRecap.Items.Add ('erreur : '+e.message);
            RollBack ;
          end;
    end;
    BeginTrans;
  end;
end;
 *)

{
procedure TFRecupSuspect.Ecrire_Datas(NomChampTable: string; NumTable,
  NumeroChamp: integer; ValeurSeq: variant; var SQL_Champs, SQL_Datas,
  SQL_Sep: string);
var
  st_trav1, st_trav2 : string;
begin
  st_trav1 := V_PGI.DEChamps[NumTable, NumeroChamp].Tipe;
  if (st_trav1='INTEGER') or (st_trav1='SMALLINT') then st_trav2 := IntToStr(ValeurSeq)    else
  if (st_trav1='DOUBLE')  or (st_trav1='RATE')     then st_trav2 := FloatToStr(ValeurSeq)  else
  if (st_trav1='DATE')                             then st_trav2 := UsDateTime(ValeurSeq)  else
                                                        st_trav2 := ValeurSeq;
  while Pos(',', st_trav2) <> 0 do st_trav2[Pos(',', st_trav2)] := '.';

  if (st_trav1 = 'DATE') or (st_trav1 = 'COMBO') or (st_trav1 = 'BOOLEAN')or (st_trav1 = 'BLOB') then
  SQL_Datas  := SQL_Datas + SQL_Sep + '"' + st_trav2 + '"'
  else if (Pos('VARCHAR', st_trav1) <> 0) then
      begin
      while Pos('"', st_trav2) <> 0 do st_trav2[Pos('"', st_trav2)] := '''';
      ReadTokenPipe(st_trav1, '(');
      st_trav1 := Copy(st_trav1, 0, Pos(')', st_trav1) - 1);
      st_trav2 := Copy(st_trav2, 0, StrToInt(st_trav1));
      SQL_Datas  := SQL_Datas + SQL_Sep + '"' + st_trav2 + '"'
      end
      else
      SQL_Datas  := SQL_Datas + SQL_Sep + st_trav2;

  SQL_Champs := SQL_Champs + SQL_Sep + V_PGI.DEChamps[NumTable, NumeroChamp].Nom;
  SQL_Sep    := ', ';
end;
}
procedure TFRecupSuspect.TraiteChamp (IndexTobTable : integer; ValChamp, Champ : string);
var TypeChamp : string;
    NumeroChamp : Integer;
    TobImport : tob;
    TraiteSpe : boolean;
    PosChamp : integer;
    PosMaxi : integer;
    NomDebut : boolean;
		Mcd : IMCDServiceCOM;

begin
NomDebut := False;
if pos ('*', champ) > 0 then TraiteSpe := true else TraiteSpe := False; //si champ de traitement spécial
if pos ('RSC_', Champ) > 0 then  TobImport := tobSuspectcompl else TobImport := TobSuspect;
if TraiteSpe then
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
   // gestion d'une raison sociale longue
   if Champ = '*Raison sociale*' then
   begin
      TypeChamp := mcd.getField('RSU_LIBELLE').tipe;
      PosMaxi := TipeToLength(TypeChamp);

      if Length(ValChamp) > PosMaxi then
      begin
         PosChamp := PosMaxi;
         while ValChamp[PosChamp] <> ' ' do
         begin
            dec (PosChamp);
            if PosChamp = 0 then
            begin
               PosChamp := PosMaxi;
               break;
            end;
         end;
         TobImport.PutValue('RSU_LIBELLE', trim(copy(ValChamp, 0, PosChamp)));
         if TobImport.GetValue('RSU_PRENOM')='' then  //si le prenom est déjà rempli dans le cas du particulier on le laisse
           TobImport.PutValue('RSU_PRENOM', trim(Copy(ValChamp, PosChamp+1, length(ValChamp))));  //remplacement de adresse1 par prenom
         exit;
         end
      else
         Champ := 'RSU_LIBELLE';
   end;

   // gestion de l'adresse1 longue
   if Champ = '*Adresse*' then
   begin
      TypeChamp := mcd.getField('RSU_ADRESSE1').tipe;
      PosMaxi := TipeToLength(TypeChamp);
      if Length(ValChamp) > PosMaxi then
      begin
         PosChamp := PosMaxi;
         while ValChamp[PosChamp] <> ' ' do
         begin
            dec (PosChamp);
            if PosChamp = 0 then
            begin
               PosChamp := PosMaxi;
               break;
            end;
         end;
            TobImport.PutValue('RSU_ADRESSE1',trim(copy(ValChamp, 0, PosChamp)));
            TobImport.PutValue('RSU_ADRESSE2', trim(Copy(ValChamp, PosChamp+1, length(ValChamp))));
      end
      else
      begin
            Champ := 'RSU_ADRESSE1';
      end;
   end;

   // gestion du nom prénom dans un même champ
   if (Champ = '*Prénom Nom*') or (Champ = '*Nom Prénom*') then
   begin
      if Champ = '*Nom Prénom*' then NomDebut := True;
      PosChamp := pos (' ', ValChamp);
      if NomDebut then
      begin
         TobImport.PutValue('RSU_CONTACTNOM', copy(ValChamp, 0, PosChamp-1));
         TobImport.PutValue('RSU_CONTACTPRENOM', copy(ValChamp, PosChamp+1, length(ValChamp)));
      end
      else
      begin
         TobImport.PutValue('RSU_CONTACTPRENOM', copy(ValChamp, 0, PosChamp-1));
         TobImport.PutValue('RSU_CONTACTNOM', copy(ValChamp, PosChamp+1, length(ValChamp)));
      end;
   exit;
   end;
end;

TypeChamp := mcd.getField(Champ).tipe;
if TypeChamp = 'DATE' then
  begin
  if IsValidDate(ValChamp) then
    TobImport.PutValue (Champ, strtodate (ValChamp));
  end
else if TypeChamp = 'INTEGER' then
  begin
  if isnumeric(ValChamp) then
    TobImport.PutValue (Champ, strtoint (ValChamp));
  end
else if (TypeChamp = 'BOOLEAN') and (trim(ValChamp) <>'') then
  begin
  if (ValChamp[1]<>'-') and (uppercase(ValChamp)<>'FAUX') and (uppercase(ValChamp)<>'FALSE') then
    TobImport.PutValue (Champ,'X');
  end
else if TypeChamp = 'DOUBLE' then
  begin
  if isnumeric(ValChamp) then
    TobImport.PutValue (Champ, valeur (ValChamp));
  end
else if (TypeChamp = 'COMBO') or (Copy(Champ, 1, 8) = 'RSC_6RSC') then
  begin
  TobImport.PutValue (Champ, copy(ValChamp,0,3));
  // quoi qu'il arrive, on stock la valeur, mais si le libelle long correspond, alors prend le bon code
//  if length(ValChamp) > 3 then
  if Trim(ValChamp) <> '' then
     GereCombo(Champ, ValChamp, Tobimport);
  end
else if (Champ = 'RSU_SIRET') then
  begin
  // spécif KPMG suppression du tiret dans le code Siret

//  ValChamp := stringreplace(ValChamp,'-','',[rfReplaceAll]);
//  ValChamp := copy(ValChamp,0,14);
  if VerifSiret(ValChamp) then
    TobImport.PutValue (Champ,ValChamp)
  else inc(Nb_ErreurSiret);
  end
else TobImport.PutValue (Champ, ValChamp);
end;

procedure TFRecupSuspect.RecupereValChamp (TobPar : TOB; Index : integer; var ValChamp : string);
var Longueur, Offset : integer;
begin
if LongFixe then
    begin
    Longueur := TobPar.GetValue ('RRL_LONGUEUR');
    Offset := TobPar.GetValue ('RRL_OFFSET');
    ValChamp := Copy (stEnreg, Offset, Longueur);
    end else
    begin
    RecupValChampEnregVar (ValChamp, TobPar.GetValue ('RRL_OFFSET')-1);
    end;
ValChamp := Trim (ValChamp);
end;

procedure TFRecupSuspect.RecupValChampEnregVar (var ValChamp : string; IndexChamp : integer);
var stEnregTrav, stSeparateur : string;
    Index, Position : integer;
begin
stSeparateur := TobParSuspect.GetValue ('RSS_SEPARATEUR');
// pb TAB
if stSeparateur = 'TAB' then
  stSeparateur := #9;
  
if stSeparateur = 'AUT' then stSeparateur := TobParSuspect.GetValue ('RSS_SEPTEXTE');
Position := 1;
stEnregTrav := stEnreg;
for Index := 0 to IndexChamp - 1 do
    begin
    Position := Position + Pos (stSeparateur, stEnregTrav);
    stEnregTrav := Copy (stEnreg, Position, Length (stEnreg) - Position + 1);
    end;
if Pos (stSeparateur, stEnregTrav) > 0 then
  ValChamp := Copy (stEnreg, Position, Pos (stSeparateur, stEnregTrav) - 1)
else
  ValChamp := stEnregTrav;
if ValChamp<>'' then
  if (ValChamp[1]='"') and (ValChamp[Length(ValChamp)]='"') then
    ValChamp := copy (valchamp,2,Length(ValChamp)-2);
end;

function TFRecupSuspect.InitProgression: longint;
var  FLigSusp :textfile;
     st : string;
     nb_lig : longint;
begin
  AssignFile(FLigSusp, RSS_FICHIER.Text);
  Reset (FLigSusp);
  if CBPremiereLigne.checked then readln (FLigSusp,st);
  nb_lig := 0;
  while ( not EOF(FLigSusp)) do
  begin
    readln (FLigSusp,st);
    inc(nb_lig);
  end;
  CloseFile(FLigSusp);
  result := nb_lig;
end;


function TFRecupSuspect.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
begin
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    begin
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    end else St_Chaine := St;
Result := St_Chaine + ' : ';
end;

procedure TFRecupSuspect.RenseigneJournalEvenement (ioerr : TIOErr);
var TOBJnal : TOB ;
    NumEvt : integer ;
    QQ : TQuery ;
begin
NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'RSU');
TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
if BStop = True then
    begin
    TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
    ListRecap.Items.add (TraduireMemoire('Interrompue par l''utilisateur'));
    end else
    begin
    if ioerr = oeOk then
        begin
        TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
        ListRecap.Items.Add(TraduireMemoire('Traitement terminé'));
        end else
        begin
        TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
        ListRecap.Items.Add (TraduireMemoire('Traitement non terminé'));
        end;
    end;
TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
end;

{==============================================================================================}
{========== Correspondance avec table libre CEGID pour fichier SCRL ===========================}
{==============================================================================================}
procedure TFRecupSuspect.CorrespondanceJur ;
Const MaxJur = 36 ;
Juridique : Array [1..MaxJur,1..2] of string[5] =(
('01','URL  '),
('02','SAR  '),
('03','SAR  '),
('04','SNC  '),
('05','SA   '),
('06','     '),
('07','SC   '),
('08','SCI  '),
('09','GIE  '),
('10','GIE  '),
('12','SA   '),
('13','CA   '),
('14','     '),
('28','ASS  '),
('29','URL  '),
('30','EAR  '),
('31','ADM  '),
('32','ADM  '),
('33','EAR  '),
('34','URL  '),
('35','URL  '),
('36','SA   '),
('37','     '),
('38','ASS  '),
('39','SA   '),
('40','SA   '),
('41','     '),
('42','SA   '),
('43','ASS  '),
('44','ADM  '),
('45','SAS  '),
('46','     '),
('48','ADM  '),
('49','ASS  '),
('51','     '),
('52','SA   '));

Var stCode1,stCode2: String ;
    i: Integer ;
begin
    stCode1 := trim(TobSuspect.GetValue ('RSU_JURIDIQUE'));
    StCode2 := '     ';
    For i:=1 to MaxJur do if (trim(StCode1)=Juridique[i,1]) then StCode2 := Juridique[i,2];
    TobSuspect.PutValue ('RSU_JURIDIQUE', StCode2);
end;

procedure TFRecupSuspect.CorrespondanceFonc ;
Const MaxFonc = 88 ;
Fonction : Array [1..MaxFonc,1..2] of string[3] =(
('000','   '),('010','PDG'),('011','   '),('020','GER'),('024','PDG'),('025','AST'),('030','DIR'),
('031','DIR'),('040','DG '),('050','DIR'),('051','DIR'),('055','AST'),('059','DIR'),('060','SG '),
('065','DAF'),('070','RA '),('080','DF '),('085','CDG'),('090','COM'),('100','DIC'),('110','DIR'),
('115','AST'),('120','DIR'),('130','DIR'),('132','DIR'),('140','DIR'),('150','RAC'),('160','DIR'),
('165','DIR'),('170','RFO'),('180','DIR'),('190','DT '),('195','DIR'),('200','DIR'),('210','DSI'),
('211','RI '),('220','REX'),('230','   '),('240','DIR'),('250','DIR'),('260','RP '),('261','RA '),
('280','DIR'),('290','DIR'),('320','DIR'),('330','DIR'),('360','DIR'),('380','DIR'),('390','DIR'),
('400','DIR'),('415','DIR'),('430','DIR'),('440','DIR'),('603','DIR'),('606','DIR'),('609','DIR'),
('612','DIR'),('615','DIR'),('618','DIR'),('621','DF '),('624','DIR'),('627','DF '),('630','DIR'),
('633','DIR'),('636','   '),('639','DIR'),('642','DIR'),('645','DIR'),('648','DIR'),('651','DIR'),
('654','DIR'),('657','DIR'),('660','DIR'),('663','DIR'),('666','DIR'),('669','DIR'),('672','DIR'),
('675','DIR'),('678','RA '),('681','DIR'),('684','DIR'),('687','DIR'),('690','DIR'),('700','RA '),
('710','RA '),('850','DIR'),('852','RA '),('988','   '));
Var stCode1,stCode2: String ;
    i: Integer ;
begin
    stCode1 := trim(TobSuspect.GetValue ('RSU_CONTACTFONCTION'));
    StCode2 := '   ';
    For i:=1 to MaxFonc do if (trim(StCode1)=Fonction[i,1]) then StCode2 := Fonction[i,2];
    TobSuspect.PutValue ('RSU_CONTACTFONCTION', StCode2);
end;

{
Tranche de chiffre d'affaire
000	- de 1 MF
001	1 à 3 MF
003	3 à 10 MF
010	10 à 20 MF
020	20 à 40 MF
050	40 à 50 MF
051	0 à 50 MF
100	50 à 100 MF
200	100 à 200 MF
300	200 à 300 MF
400	300 à 400 MF
500	400 à 500 MF
700	500 à 700 MF
999	supèrieur a 700 MF
}

procedure TFRecupSuspect.CorrespondanceCA ;
var  montant_CA : double;
     stcode : string;
begin
  montant_CA := TobSuspectcompl.GetValue ('RSC_RSCLIBVAL1')/1000000;
  if   (montant_CA < 1) then stcode := '000'
  else if ((montant_CA >= 1) and (montant_CA < 3)) then stcode := '001'
   else if ((montant_CA >= 3) and (montant_CA < 10)) then stcode := '003'
    else if ((montant_CA >= 10) and (montant_CA < 20)) then stcode := '010'
     else if ((montant_CA >= 20) and (montant_CA < 40)) then stcode := '020'
      else if ((montant_CA >= 40) and (montant_CA < 50)) then stcode := '050'
       else if ((montant_CA >= 0) and (montant_CA < 50)) then stcode := '051'
        else if ((montant_CA >= 50) and (montant_CA < 100)) then stcode := '100'
         else if ((montant_CA >= 100) and (montant_CA < 200)) then stcode := '200'
          else if ((montant_CA >= 200) and (montant_CA < 300)) then stcode := '300'
           else if ((montant_CA >= 300) and (montant_CA < 400)) then stcode := '400'
            else if ((montant_CA >= 500) and (montant_CA < 700)) then stcode := '700'
            else if (montant_CA >= 700) then stcode := '999';
  TobSuspectcompl.PutValue ('RSC_RSCLIBTABLE0',stcode);
end;

{
Tranche de Nombre d'effectifs
001	0 à 50 personnes
005	1 à 5 personnes
010	6 à 10 personnes
011	11 à 20 personnes
020	21 à 50 personnes
050	51 à 100 personnes
100	101 à 200 personnes
200	201 à 500 personnes
500	501 à 1000 personnes
996	supèrieur à 1000 personnes
}
procedure TFRecupSuspect.CorrespondanceEff ;
var  nombre_Eff : double;
     stcode : string;
begin
  nombre_Eff := TobSuspectcompl.GetValue ('RSC_RSCLIBVAL2');
 if ((nombre_Eff >= 1) and (nombre_Eff <= 5)) then stcode := '005'
    else if ((nombre_Eff >= 6) and (nombre_Eff <= 10)) then stcode := '010'
     else if ((nombre_Eff >= 11) and (nombre_Eff <= 20)) then stcode := '011'
      else if ((nombre_Eff >= 21) and (nombre_Eff <= 50)) then stcode := '020'
       else if ((nombre_Eff >= 51) and (nombre_Eff <= 100)) then stcode := '050'
         else if ((nombre_Eff >= 101) and (nombre_Eff <= 200)) then stcode := '100'
          else if ((nombre_Eff >= 201) and (nombre_Eff <= 500)) then stcode := '200'
           else if ((nombre_Eff >= 501) and (nombre_Eff <= 1000)) then stcode := '500'
            else if (nombre_Eff >= 1000) then stcode := '996';
  TobSuspectcompl.PutValue ('RSC_RSCLIBTABLE1',stcode);
end;

procedure TFRecupSuspect.CorrespondanceDateCreation ;
var     stmois,stannee,datecreation : string;
begin
  stmois := TobSuspectcompl.GetValue ('RSC_RSCLIBTEXTE2');
  stannee := TobSuspectcompl.GetValue ('RSC_RSCLIBTEXTE3');
  if (trim(stannee) <> '') then
  begin
    if (trim(stmois) <> '') then
      datecreation := '01/'+ stmois + '/' + stannee
      else datecreation := '01/01/'+ stannee;
    if IsValidDate(datecreation) then
          TobSuspectcompl.PutValue ('RSC_RSCLIBDATE0', strtodate (datecreation));
  end;
end;
{==========================Fin correspondance table libre CEGID pour fichier SCRL ==========}

function TFRecupSuspect.VerifAccesFichier : boolean;
var Fichier : textfile;
begin
Result := True;
AssignFile (Fichier, RSS_FICHIER.Text);
{$i-}
Reset (Fichier);
{$i+}
if IoResult<>0 then
begin
  Msg.Execute(11, Caption, '');
  Result := False;
end
else CloseFile(Fichier);
end;

procedure AGLRTEntreeRecupSuspect( parms: array of variant; nb: integer ) ;
begin
  if (string(parms[0]) <> '') then
  EntreeRecupSuspect(string(parms[0]))
  else
  EntreeTrtRecupSuspect(taCreat,'','');
end;

procedure AGLRTSupRecupSuspect( parms: array of variant; nb: integer ) ;
begin
  if (string(parms[0]) <> '') then
    SupRecupSuspect(string(parms[0]));
end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 25/07/2006
Modifié le ... :   /  /
Description .. : Parcours la tablette d'un combo afin de retrouver à partir du
Suite ........ : libelle long le bon code
Mots clefs ... : TABLETTE; COMBO;
*****************************************************************}
procedure TFRecupSuspect.GereCombo(Champ, ValChamp: String; Latob : TOB);
var
//   LaTable : string;
//   TipeTablette : String;
   TobC : TOB;
   TobT : TOB;
   {
   Stt : HTstrings;
   i : integer;
   VPDE : TDECombo;
   ValCombo : string;
   }

begin
//  LaTable := ChampToTT(champ);
//  TipeTablette := TTToTipe (LaTable);

  TobC := TobCorresp.FindFirst(['CC_LIBRE', 'CC_LIBELLE'], [Champ, Copy(ValChamp, 1, 35)], False);
  If TobC <> nil then   // le libellé est dans la table de correspondance
  begin
    if TobC.GetValue('CC_ABREGE') = '' then
      BoAjoutCorr := True;
//      ListRecap.Items.Add('Le champ <'+Champ+'> pour la valeur "'+ValChamp+'" n''a pas de code');

    Latob.PutValue (Champ, TobC.GetValue('CC_ABREGE'));
  end
  else
  begin
    TobC := TobTablette.FindFirst(['CHAMP', 'LIBELLE'], [Champ, Copy(ValChamp, 1, 35)], False);
    if TobC <> nil then // le libellé est dans la tablette associée
    begin
      Latob.PutValue (Champ, TobC.GetValue('CODE'));

      TobT := Tob.Create('CHOIXCOD', TobCorresp, -1);
      TobT.PutValue ('CC_TYPE', 'CIM');
      TobT.PutValue ('CC_CODE', ProchainCodeLibre);
      TobT.PutValue ('CC_LIBELLE', Copy(ValChamp, 1, 35));
      TobT.PutValue ('CC_ABREGE', TobC.GetValue('CODE'));
      TobT.PutValue ('CC_LIBRE', Champ);
      ListRecap.Items.Add ('Ajout d''un lien sur <'+Champ+'> avec le libellé '+ValChamp);
      BoAjoutCorr := True;
      ChercheProchainCode;
    end
    else
    begin            // Le libellé est introuvable. on créé quand même la correspondance
      TobT := Tob.Create('CHOIXCOD', TobCorresp, -1);
      TobT.PutValue ('CC_TYPE', 'CIM');
      TobT.PutValue ('CC_CODE', ProchainCodeLibre);
      TobT.PutValue ('CC_LIBELLE', Copy(ValChamp, 1, 35));
      TobT.PutValue ('CC_ABREGE', '');
      TobT.PutValue ('CC_LIBRE', Champ);
      ListRecap.Items.Add ('Ajout d''une correspondance pour le champ <'+Champ+'> sur le libellé "'+ValChamp+'"');
      BoAjoutCorr := True;
      ChercheProchainCode;
    end;
  end;


{
   LaTable := ChampToTT(champ);

   i := TTToNum(LaTable);
   if i <=0 then
      exit;

 //  RemplirListe(LaTable,'');
   VPDE := V_PGI.DECombos[i];
   Stt := VPDE.Valeurs;
   if (Stt = nil) or (Stt.Count-1 = 0) then
      exit;
   for i:=0 to Stt.Count-1 do
   begin
      ValCombo := UpperCase_(ExtractInfo(Stt[i], 1));
      if  ValCombo = Copy(UpperCase_(ValChamp),0, length(ValCombo)) then
      begin
         Latob.PutValue(Champ, ExtractInfo(Stt[i], 0));
         exit;
      end;
   end;
}
end;

procedure TFRecupSuspect.CBControleProspectClick(Sender: TObject);
begin
  inherited;
   if CBControleProspect.Checked then
      CBMajTiers.Enabled := True
   else
   begin
      CBMajTiers.Checked := False;
      CBMajTiers.Enabled := False;
   end;
end;

procedure TFRecupSuspect.CBMajTiersClick(Sender: TObject);
begin
  inherited;
  if CBMajTiers.Checked then
  begin
      GBTiers.Enabled := True;
      CbProspect.Enabled  := True;
      CbClient.Enabled    := True;
      CbAutre.Enabled     := True;
      CbProspect.Checked  := True;
      CbClient.Checked    := True;
      CbAutre.Checked     := True;
      CBContact.Enabled := True;
      CBContact.Checked := True;
  end
  else
  begin
      GBTiers.Enabled := False;
      CbProspect.Enabled  := False;
      CbClient.Enabled    := False;
      CbAutre.Enabled     := False;
      CBContact.Checked := False;
      CBContact.Enabled := False;
  end;

end;

procedure TFRecupSuspect.ChargeTablette;
var
  TobImport : TOB;
  NumChamp : Integer;
  TypeChamp : String;
  LaTable : String;
  i : integer;
  j : integer;
  Champ : String;
  Stt : HTstrings;
  k : integer;
  VPDE : TDECombo;
  TobT : TOB;
  TobParsuspectCor      : Tob;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  TobTablette := Tob.Create('Les Tablettes', nil, -1);

  for i := 0 to TobParSuspectLig.Detail.Count - 1 do
    begin
      Champ := TobParSuspectLig.Detail[i].GetValue ('RRL_CHAMP');
      if pos ('*',Champ) = 1 then
        Continue;
      if pos ('RSC_',Champ) >0 then
        TobImport := TobSuspectCompl
      else
        TobImport := TobSuspect;

      //on teste les champs qui vont être importés et s'il s'agit de table on les charge
      NumChamp := TobImport.GetNumChamp(Champ);
      TypeChamp := mcd.getField(Champ).tipe;
      If (Typechamp = 'COMBO') or (Copy(Champ, 1, 8) = 'RSC_6RSC')  then
      begin
        //on regarde si il y a une correspondance avec une tablette tiers
        LaTable         := RenvoiTabletteCor(TobParsuspectCor, Champ);
        //sinon on prend celle liée au champs
        if LaTable = '' then
          LaTable := ChampToTT(Champ);

        j := TTToNum(LaTable);
        if j<=0 then
          Continue;
        RemplirListe(LaTable, '');
        VPDE := V_PGI.DECombos[j];
        Stt := VPDE.Valeurs;

        for k := 0 to Stt.Count -1 do
        begin
          TobT := Tob.Create ('Tablette', TobTablette, -1);
          TobT.AddChampSupValeur ('CHAMP', Champ);
          TobT.AddChampSupValeur ('CODE', ExtractInfo(Stt[k], 0));
          TobT.AddChampSupValeur ('LIBELLE', ExtractInfo(Stt[k], 1));
        end;

      end;
    end;
    FreeAndNil(TobParsuspectCor);
end;




function TFRecupSuspect.TraduireFormule(LaFormule, LeChamp : String): String;
var
  Pos1 : integer;
  Pos2 : integer;
  Prefixe : String;
  Itable : integer;
  I : integer;
  LeType : String;
  StrRes : String;
  ChampRemp : String;
  OffSet : String;
  Longeur : String;
  ValChamp : String;
	Mcd : IMCDServiceCOM;

begin

try

  Result := '';
	Letype := mcd.getField(LeChamp).tipe;
  Pos1 := Pos('{', LaFormule);
  While Pos1 <> 0 do
  begin
    Pos2 := Pos('}', LaFormule);
    if Pos2 < Pos1 then
    begin
      Break;
    end;
    ChampRemp := Copy(LaFormule, Pos1, Pos2-Pos1+1);
    if LongFixe then
    begin
      Pos2 := Pos('-', ChampRemp);
      OffSet := Copy(ChampRemp, 2, Pos2-2);
      Longeur := Copy(ChampRemp, pos2+1, length(ChampRemp)-Pos2-1);
      ValChamp := Copy(stEnreg, StrToInt(OffSet), StrToInt(Longeur));
    end
    else
    begin
      OffSet := Copy(ChampRemp, 2, Length(ChampRemp)-2);
      RecupValChampEnregVar(ValChamp, StrToInt(OffSet)-1);
    end;
    if (Letype = 'DOUBLE') or (LeType = 'INTEGER') then
      LaFormule := StringReplace_(LaFormule , ChampRemp, ValChamp, [rfReplaceAll])
    else
      LaFormule := StringReplace_(LaFormule , ChampRemp, '("'+ValChamp+'")', [rfReplaceAll]);

    Pos1 := Pos('{', LaFormule);
  end;

  except
     on e:exception do
     begin
        // rajout dans le log
        ListRecap.Items.Add ('erreur : '+e.Message+' -> '+LaFormule+','+LeChamp  );
        StrRes := '' ;
     end;
  end;



  try
     StrRes := TheFormVM.GetExprValue(LaFormule);

  except
     on e:exception do
     begin
        // rajout dans le log
        ListRecap.Items.Add ('erreur : '+e.Message+' -> '+LaFormule+','+LeChamp  );
        StrRes := '' ;
     end;
  end;

  Result := StrRes;

end;



procedure TFRecupSuspect.ChargeTableCorresp;
var
  Q : TQuery;

begin
  TobCorresp := Tob.Create('Les correspondances', nil, -1);
  Q := OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CIM"', True,-1,'',true);
  TobCorresp.LoadDetailDB('CHOIXCOD', '', '',Q,True);
  Ferme(Q);
end;



procedure TFRecupSuspect.ChercheProchainCode;
var
  NumCode : Integer;

begin

  if TobCorresp.Detail.Count = 0 then
    ProchainCodeLibre := '001'
  else
  begin
    ProchainCodeLibre := TobCorresp.Detail[TobCorresp.Detail.Count -1].GetValue('CC_CODE');
    NumCode := StrToInt(ProchainCodeLibre);
    Inc(NumCode);
    ProchainCodeLibre := IntToStr(NumCode);
    if Length (ProchainCodeLibre) = 1 then
      ProchainCodeLibre := '00' + ProchainCodeLibre;
    if Length (ProchainCodeLibre) = 2 then
      ProchainCodeLibre := '0' + ProchainCodeLibre;
  end;

end;

procedure TFRecupSuspect.CBtestClick(Sender: TObject);
begin
  inherited;
  if CBtest.Checked then
  begin
    CBCreationSuspect.Enabled := False;
    CBCreationSuspect.Checked := False;
    CBControleProspect.Enabled := false;
    CBControleProspect.Checked := false;
    CBMajTiers.Enabled := false;
    CBMajTiers.Checked := False;
  end
  else
  begin
    CBCreationSuspect.Enabled := True;
    CBCreationSuspect.Checked := False;
    CBControleProspect.Enabled := True;
    CBControleProspect.Checked := false;
//    CBMajTiers.Enabled := true;
//    CBMajTiers.Checked := False;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJA
Créé le ...... : 08/08/2007
Modifié le ... :   /  /    
Description .. : Mise à jour des tables TIERS, TIERSCOMPL, PROSPECTS 
Suite ........ : Le fichier en entrée doit impérativement avoir T_TIERS.
Mots clefs ... : 
*****************************************************************}
procedure TFRecupSuspect.MajTiersParCode;
var
  Fichier               : textfile;
  NligneFichier         : longint;
  i                     : integer;
  ValChamp              : String;
  Champ                 : String;
  LaFormule             : String;
  NbMaj                 : integer;
  NbLue                 : integer;
  Broll                 : boolean;

begin

  NumchampCodeTiers     := -1;
  for i := 0 to TobParSuspectLig.Detail.Count -1 do
  begin
    if TobParSuspectLig.Detail[i].GetValue('RRL_CHAMP') = 'T_TIERS' then
    begin
      NumchampCodeTiers := i;
      break;
    end;
  end;
  if NumchampCodeTiers = -1 then
  begin
    PGIInfo('Il manque le champ ''code tiers'' dans la decription du fichier', PTITRE.Caption);
    exit;
  end;

  NligneFichier         := InitProgression ;
  ListRecap.Items.Add('Nb éléments à traiter : ' + IntToStr(NligneFichier));

  InitMoveProgressForm (nil,TraduireMemoire('Traitement en cours...'), '', NligneFichier, True, True) ;

  AssignFile (Fichier, RSS_FICHIER.Text); // ouvre le pointeur sur le fichier
  Reset(Fichier);
  if CBPremiereLigne.checked then
    readln (Fichier, stEnreg);

  Bstop                 := not(MoveCurProgressForm (TraduireMemoire('Mise à jour des fiches tiers')));
  NbMaj                 := 0;
  NbLue                 := 0;
  Broll                 := False;

  while not EOF(Fichier) and not Bstop do
  begin
    if Not MoveCurProgressForm(TraduireMemoire('Mise à jour des fiches tiers')) then
      break ;
    Readln(Fichier,stEnreg);
    try
      if (NbMaj mod KBlockTiers = 0) then
        BEGINTRANS;
      inc(NbLue);
      RecupereValChamp(TobParSuspectLig.Detail[NumchampCodeTiers], NumchampCodeTiers, Valchamp);
      if ExisteTiers(ValChamp) then
      begin
        for i := 0 to TobParSuspectLig.Detail.Count -1 do
        begin
          // si code tiers, on passe
          if i = NumchampCodeTiers then
            Continue;

          Champ           := TobParSuspectLig.Detail[i].GetValue('RRL_CHAMP');

          RecupereValChamp(TobParSuspectLig.Detail[i], i, Valchamp);
          LaFormule       := TobParSuspectLig.Detail[i].GetValue('RRL_FORMULE');
          if Trim(LaFormule) <> '' then
            ValChamp      := TraduireFormule(LaFormule, Champ);
          TraiteChampCodeTiers(Champ, ValChamp);
        end;
        EnregLesTiers;
        inc(NbMaj);
      end;

      if (NbMaj mod KBlockTiers = 0) then
        COMMITTRANS;

    Except
      on E:Exception do
      begin
        ROLLBACK;
        if V_PGI.IOError = oeReseau then
          ListRecap.Items.Add('Problème réseau');
        ListRecap.Items.Add('Arrêt sur le ' + IntToStr(NbLue) + 'ème enregistrement');
        ListRecap.Items.Add('Code tiers : ' + TobTiers.GetValue('T_TIERS'));
        Broll           := True;
      end;
    end;
  end;

  // au cas où le block d'enregistrement n'est pas complet
  try
    if (NbMaj mod KBlockTiers <> 0) then
      COMMITTRANS;
  except
    on E:Exception do
    begin
      ROLLBACK;
    end;
  end;

  if not Broll then
  begin
    ListRecap.Items.Add('Nb Fiches tiers modifiées : ' + IntToStr(NbMaj));
    ListRecap.Items.Add('Mise à jour terminée');
  end
  else
  begin
    ListRecap.Items.add('Mise à jour interrompue !');
  end;

  CloseFile(Fichier);
  FiniMoveProgressForm;
  RenseigneJournalEvenement(oeOk);

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJA
Créé le ...... : 08/08/2007
Modifié le ... :   /  /    
Description .. : Vérifie l'existance d'une fiche tiers à partir du code T_TIERS
Mots clefs ... : 
*****************************************************************}
function TFRecupSuspect.ExisteTiers(CodeTiers: String): boolean;
var
  Q                     : TQuery;

begin
  Result                := False;
  TobTiers              := tob.Create('TIERS', nil, -1);
  Q                     := OpenSQL('SELECT * FROM TIERS WHERE T_TIERS = "' + CodeTiers + '"', True,-1,'',true);
  if not Q.Eof then
  begin
    TobTiers            := Tob.Create('TIERS', nil, -1);
    Tobtiers.SelectDB('', Q);
    TobYtc              := Tob.Create('TIERSCOMPL', nil, -1);
    TobRpr              := Tob.Create('PROSPECTS', nil, -1);
    TobYtc.SelectDB('"' + TobTiers.GetValue('T_AUXILIAIRE') + '"', nil);
    TobRpr.SelectDB('"' + TobTiers.GetValue('T_AUXILIAIRE') + '"', nil);
    //au cas où les tables compléments ne sont renseignées
    TobYtc.PutValue('YTC_AUXILIAIRE', TobTiers.GetValue('T_AUXILIAIRE'));
    TobYtc.PutValue('YTC_TIERS', TobTiers.GetValue('T_TIERS'));
    TobRpr.PutValue('RPR_AUXILIAIRE', TobTiers.GetValue('T_AUXILIAIRE'));
    Result              := True;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJA
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : Mise à jour des différentes Tob
Mots clefs ... : 
*****************************************************************}
procedure TFRecupSuspect.TraiteChampCodeTiers(LeChamp, LeValChamp: String);
var
  NumeroChamp           : integer;
  TypeChamp             : String;
  TobM                  : Tob;
	Mcd : IMCDServiceCOM;

begin
  if copy(LeChamp, 1, 2) = 'T_' then
    TobM                := TobTiers
  else if copy(LeChamp, 1, 4) = 'YTC_' then
    TobM                := TobYtc
  else
    TobM                := tobRpr;

  TypeChamp:= mcd.getField(LeChamp).tipe;
  if TypeChamp = 'DATE' then
    begin
    if IsValidDate(LeValChamp) then
      TobM.PutValue (LeChamp, strtodate (LeValChamp));
    end
  else if TypeChamp = 'INTEGER' then
    begin
    if isnumeric(LeValChamp) then
      TobM.PutValue (LeChamp, strtoint (LeValChamp));
    end
  else if (TypeChamp = 'BOOLEAN') and (trim(LeValChamp) <>'') then
    begin
    if (LeValChamp[1]<>'-') and (uppercase(LeValChamp)<>'FAUX') and (uppercase(LeValChamp)<>'FALSE') then
      TobM.PutValue (LeChamp,'X');
    end
  else if TypeChamp = 'DOUBLE' then
    begin
    if isnumeric(LeValChamp) then
      TobM.PutValue (LeChamp, valeur (LeValChamp));
    end
  else if TypeChamp = 'COMBO' then
    begin
    TobM.PutValue (LeChamp, LeValChamp);
//    if Trim(ValChamp) <> '' then
//       GereCombo(Champ, ValChamp, Tobimport);
    end
  else TobM.PutValue (LeChamp, LeValChamp);

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJA
Créé le ...... : 08/08/2007
Modifié le ... :   /  /
Description .. : update ou création (TiersCompl et/ou Prospects)
Mots clefs ... :
*****************************************************************}
procedure TFRecupSuspect.EnregLesTiers;
begin
  TobTiers.SetAllModifie(True);
  TobYtc.SetAllModifie(True);
  TobRpr.SetAllModifie(True);

  TobTiers.UpdateDB(False);
  TobYtc.InsertOrUpdateDB(False);
  TobRpr.InsertOrUpdateDB(False);

end;


procedure TFRecupSuspect.ListRecapMajCodeTiers;
begin
  ListRecap.Clear;
  ListRecap.Items.Add('Mise à jour des fiches Tiers');
  ListRecap.Items.Add('----------------------------');
  ListRecap.Items.Add('Fichier : ' + RSS_FICHIER.Text);
  if CBPremiereLigne.Checked then
    ListRecap.Items.Add('Ligne d''entête supprimée');

end;

Initialization
RegisterAglProc( 'RecupSuspect', FALSE , 1, AGLRTEntreeRecupSuspect);
RegisterAglProc( 'SupRecupSuspect', FALSE , 1, AGLRTSupRecupSuspect);
end.
