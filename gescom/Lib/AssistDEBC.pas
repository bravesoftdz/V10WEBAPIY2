unit AssistDEBC;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  HPanel, Mask, UTOB, HEnt1, DBTables, UtilGC, Ent1, UIUtil, Ed_Tools,
  EntGC,ParamSoc,HStatus,HXlsPas;


type
  TFAssistDEBC = class(TFAssist)
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    PTITRE: THPanel;
    PBevel1: TBevel;
    TINTRO: THLabel;
    GBOPTIONS: TGroupBox;
    CB_CLIENTS: TCheckBox;
    CB_ARTICLES: TCheckBox;
    CB_EXPORT: TCheckBox;
    CB_INTRO: TCheckBox;
    TEXPLIQ: THLabel;
    TabSheet6: TTabSheet;
    PBevel2: TBevel;
    Bevel1: TBevel;
    GBCLIENTS: TGroupBox;
    CDATECLI: THCritMaskEdit;
    TDATECLI: THLabel;
    TFICCLI: THLabel;
    NomFicCli: TEdit;
    RechFicCli: TToolbarButton97;
    GBARTICLES: TGroupBox;
    TDATEART: THLabel;
    TFICART: THLabel;
    RechFicArt: TToolbarButton97;
    CDATEART: THCritMaskEdit;
    NomFicArt: TEdit;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    LocaliseArt: TSaveDialog;
    LocaliseCli: TSaveDialog;
    GBEXPORT: TGroupBox;
    TDATEDE: THLabel;
    DateDebExp: THCritMaskEdit;
    TDATEFE: THLabel;
    DateFinExp: THCritMaskEdit;
    TKILOE: THLabel;
    KiloE: THValComboBox;
    TLIBREE: THLabel;
    MontantLibreE: THValComboBox;
    TFICEXP: THLabel;
    NomFicExport: TEdit;
    RechFicExport: TToolbarButton97;
    LocaliseExport: TSaveDialog;
    Bevel5: TBevel;
    GBINTRO: TGroupBox;
    TDATEDI: THLabel;
    TDATEFI: THLabel;
    TKILOI: THLabel;
    TLIBREI: THLabel;
    TFICINT: THLabel;
    RechFicIntro: TToolbarButton97;
    DateDebInt: THCritMaskEdit;
    DateFinInt: THCritMaskEdit;
    KiloI: THValComboBox;
    MontantLibreI: THValComboBox;
    NomFicIntro: TEdit;
    PBevel4: TBevel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    LocaliseIntro: TSaveDialog;
    HMsgErr: THMsgBox;
    MsgBox: THMsgBox;
    Infoscli: THLabel;
    Infosarticles: THLabel;
    InfosExport: THLabel;
    InfosIntro: THLabel;
    Bfermer: TButton;
    GroupBox2: TGroupBox;
    HLabel1: THLabel;
    FRMT_EXC: TRadioButton;
    FRMT_SDF: TRadioButton;
    procedure RechFicCliOnClick(Sender: TObject);
    procedure RechFicArtOnClick(Sender: TObject);
    procedure RechFicExportOnClick(Sender: TObject);
    procedure RechFicIntroOnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CDATECLIOnExit(Sender: TObject);
    procedure CDATEARTOnExit(Sender: TObject);
    procedure DateDebExpOnExit(Sender: TObject);
    procedure DateFinExpOnExit(Sender: TObject);
    procedure DateDebIntOnExit(Sender: TObject);
    procedure DateFinIntOnExit(Sender: TObject);
    procedure KiloEOnClick(Sender: TObject);
    procedure KiloIOnClick(Sender: TObject);
    procedure MontantLibreEOnClick(Sender: TObject);
    procedure MontantLibreIOnClick(Sender: TObject);
    procedure CB_CLIENTSOnClick(Sender: TObject);
    procedure CB_ARTICLESOnClick(Sender: TObject);
    procedure CB_EXPORTOnClick(Sender: TObject);
    procedure CB_INTROOnClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure BfermerClick(Sender: TObject);
    procedure FRMT_SDFClick(Sender: TObject);
    procedure FRMT_EXCClick(Sender: TObject);
  private
    { Déclarations privées }
    i_numecran : integer;
    AuMoinsUn, Excel : boolean;
    NbItems : integer;
    Ext1 : string;
    Procedure ListeRecap;
    Function TraitementClients : boolean;
    Function TraitementArticles : boolean;
    Function TraitementDEB(TypeDEB : string) : boolean;
    Procedure ChargerLesPieces(TypeDEB : string; var TOBGlobal : TOB; var WhereLigne : string);
    Procedure ChargerLesLignes(TypeDEB :string; var TOBGlobal : TOB; WhereLigne : string; var nblig : integer);
    Procedure ChargerLesNoms(TypeDEB : string; WhereNom : string; var TOBNomenc : TOB);
    Procedure FormatChaine(var zone : string; lg : integer);
    Procedure WhereParPiece(TypeDEB : string; var TOBNat : TOB; var WhereNat : string);
    Procedure EnregistreEvenement;
    Procedure Eclater(var TobEclat : TOB; TOBNomenc : TOB; NumLigne : integer; Niveau : Integer; Compose : string; Ordre : Integer; Qte : double);
    Procedure TobEclatToTobLigne(TypeDEB : string; var TobPiece : TOB; TobLigne : TOB; TOBCompos : TOB; var NbLig : Integer);
    Procedure FilesName;
  public
    { Déclarations publiques }
  end;

var
  FAssistDEBC: TFAssistDEBC;

Procedure Assist_Debc;

implementation

{$R *.DFM}

//========================================================
// Point d'entrée
//========================================================
Procedure Assist_Debc;
var FF : TFAssistDEBC;
begin
FF := TFAssistDEBC.Create(Application);
  try
     FF.BorderStyle := bsSingle;
     FF.ShowModal;
  finally
     FF.Free;
  end;
end;

//==========================================================================
// gestion de l'écran
//==========================================================================

procedure TFAssistDEBC.FormShow(Sender: TObject);
begin
  inherited;
  bAnnuler.Visible := True;
  bSuivant.Enabled := False;
  bFin.Visible := True;
  bFin.Enabled := False;
  bFermer.Visible := False;
  FRMT_EXC.Checked := True; // JTR DEB
  FRMT_SDF.Checked := False; // JTR DEB
  Ext1 := '.xls';
  if not FileExists ('C:\INSTAT') then CreateDir ('C:\INSTAT'); // DBR Fiche 10128
  FilesName;
  MontantLibreE.value := '<<Aucun>>';
  MontantLibreI.value := '<<Aucun>>';
end;

//=========================================
// localisation des divers fichiers
//=========================================
procedure TFAssistDEBC.RechFicCliOnClick(Sender: TObject);
begin
  inherited;
  DirDefault(LocaliseCli,NomFicCli.Text);
  If LocaliseCli.execute Then NomFicCli.text := LocaliseCli.FileName;
end;

procedure TFAssistDEBC.RechFicArtOnClick(Sender: TObject);
begin
  inherited;
  DirDefault(LocaliseArt,NomFicArt.Text);
  If LocaliseArt.execute Then NomFicArt.text := LocaliseArt.FileName;
end;

procedure TFAssistDEBC.RechFicExportOnClick(Sender: TObject);
begin
  inherited;
  DirDefault(LocaliseExport,NomFicExport.Text);
  If LocaliseExport.execute Then NomFicExport.text := LocaliseExport.FileName;
end;

procedure TFAssistDEBC.RechFicIntroOnClick(Sender: TObject);
begin
  inherited;
  DirDefault(LocaliseIntro,NomFicIntro.Text);
  If LocaliseIntro.execute Then NomFicIntro.text := LocaliseIntro.FileName;
end;

//=========================================
// controles de dates
//=========================================

procedure TFAssistDEBC.CDATECLIOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(CDATECLI.Text) then
  begin
    stErr:='"'+ CDATECLI.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    CDATECLI.SetFocus;
  end;
end;

procedure TFAssistDEBC.CDATEARTOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
   if not IsValidDate(CDATEART.Text) then
  begin
    stErr:='"'+ CDATEART.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    CDATEART.SetFocus;
  end;
end;

procedure TFAssistDEBC.DateDebExpOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DateDebExp.Text) then
  begin
    stErr:='"'+ DateDebExp.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DateDebExp.SetFocus;
  end else
  begin
    NomFicExport.text := 'C:\INSTAT\PGIE'+trim(Copy(DateDebExp.text,4,2))+'.sdf';
  end;
end;

procedure TFAssistDEBC.DateFinExpOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DateFinExp.Text) then
  begin
    stErr:='"'+ DateFinExp.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DateFinExp.SetFocus;
  end;
end;


procedure TFAssistDEBC.DateDebIntOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DateDebInt.Text) then
  begin
    stErr:='"'+ DateDebInt.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DateDebInt.SetFocus;
  end else
  begin
    NomFicIntro.text := 'C:\INSTAT\PGIE'+trim(Copy(DateDebInt.text,4,2))+'.sdf';
  end;
end;

procedure TFAssistDEBC.DateFinIntOnExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DateFinInt.Text) then
  begin
    stErr:='"'+ DateFinInt.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DateFinInt.SetFocus;
  end;
end;

//=====================================
// Unité Kilo
//=====================================
procedure TFAssistDEBC.KiloEOnClick(Sender: TObject);
begin
  inherited;
  if KiloI.value = '' then KiloI.value := KiloE.value;
end;

procedure TFAssistDEBC.KiloIOnClick(Sender: TObject);
begin
  inherited;
  if KiloE.value = '' then KiloE.value := KiloI.value;
end;

//======================================
// montants libres
//======================================

procedure TFAssistDEBC.MontantLibreEOnClick(Sender: TObject);
begin
  inherited;
  if MontantLibreI.value = '<<Aucun>>' then MontantLibreI.value := MontantLibreE.value;
end;

procedure TFAssistDEBC.MontantLibreIOnClick(Sender: TObject);
begin
  inherited;
  if MontantLibreE.value = '<<Aucun>>' then MontantLibreE.value := MontantLibreI.value;
end;

//==============================================
// gestion des onglets et des boutons prec/suiv
//===============================================

// on ne passe à l'onglet suivant que si on a choisi un fichier

procedure TFAssistDEBC.CB_CLIENTSOnClick(Sender: TObject);
begin
  inherited;
  bSuivant.Enabled:=(CB_CLIENTS.Checked) or (CB_ARTICLES.Checked)
  or (CB_EXPORT.Checked) or (CB_INTRO.Checked);
end;

procedure TFAssistDEBC.CB_ARTICLESOnClick(Sender: TObject);
begin
  inherited;
  bSuivant.Enabled:=(CB_CLIENTS.Checked) or (CB_ARTICLES.Checked)
  or (CB_EXPORT.Checked) or (CB_INTRO.Checked);
end;

procedure TFAssistDEBC.CB_EXPORTOnClick(Sender: TObject);
begin
  inherited;
  bSuivant.Enabled:=(CB_CLIENTS.Checked) or (CB_ARTICLES.Checked)
  or (CB_EXPORT.Checked) or (CB_INTRO.Checked);
end;

procedure TFAssistDEBC.CB_INTROOnClick(Sender: TObject);
begin
  inherited;
  bSuivant.Enabled:=(CB_CLIENTS.Checked) or (CB_ARTICLES.Checked)
  or (CB_EXPORT.Checked) or (CB_INTRO.Checked);
end;

procedure TFAssistDEBC.FRMT_SDFClick(Sender: TObject);
begin
  FilesName;
end;

procedure TFAssistDEBC.FRMT_EXCClick(Sender: TObject);
begin
  FilesName;
end;
// on saute les onglets non concernés (suivant)

procedure TFAssistDEBC.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
  Onglet := P.ActivePage;
  st_NomOnglet := Onglet.Name;
  i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1));
  if (CB_CLIENTS.Checked = False) and (i_NumEcran = 2) then
    BEGIN
    bSuivantClick(self);
    exit;
    END;
if (CB_ARTICLES.Checked = False) and (i_NumEcran = 3) then
    BEGIN
    bSuivantClick(self);
    exit;
    END;
if (CB_EXPORT.Checked = False) and (i_NumEcran = 4) then
    BEGIN
    bSuivantClick(self);
    exit;
    END;
if (CB_INTRO.Checked = False) and (i_NumEcran = 5) then
    BEGIN
    bSuivantClick(self);
    exit;
    END;
if i_numecran = 6 then ListeRecap;
if (bSuivant.Enabled) then bFin.Enabled := False
   else bFin.Enabled := True;
end;

// on saute les onglets non concernés (Précédent)

procedure TFAssistDEBC.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
  bFermer.Visible := False;
  bFin.Visible := True;
  Onglet := P.ActivePage;
  st_NomOnglet := Onglet.Name;
  i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1));
  if (CB_CLIENTS.Checked = False) and (i_NumEcran = 2) then
    BEGIN
    bPrecedentClick(self);
    exit;
    END;
  if (CB_ARTICLES.Checked = False) and (i_NumEcran = 3) then
   BEGIN
    bPrecedentClick(self);
    exit;
    END;
  if (CB_EXPORT.Checked = False) and (i_NumEcran = 4) then
    BEGIN
    bPrecedentClick(self);
    exit;
    END;
  if (CB_INTRO.Checked = False) and (i_NumEcran = 5) then
    BEGIN
    bPrecedentClick(self);
    exit;
    END;
if (bSuivant.Enabled) then bFin.Enabled := False
   else bFin.Enabled := True;
end;

// JT - Nom des fichiers selon choix Formats
Procedure TFAssistDEBC.FilesName;
var Ext : string;
begin
  if FRMT_EXC.Checked then
  begin
    Excel := True;
    Ext := '.xls';
  end else
  begin
    Excel := False;
    Ext := '.sdf';
  end;
  if NomFicCli.text = '' then
    NomFicCli.text := 'C:\INSTAT\PGICLI'+Ext
    else
    NomFicCli.text := Copy(NomFicCli.text,1,pos(Ext1,NomFicCli.text)-1)+Ext;
  if NomFicArt.text = '' then
    NomFicArt.text := 'C:\INSTAT\PGIART'+Ext
    else
    NomFicArt.text := Copy(NomFicArt.text,1,pos(Ext1,NomFicArt.text)-1)+Ext;
  if NomFicExport.text = '' then
    NomFicExport.text := 'C:\INSTAT\PGIE'+trim(Copy(DateDebExp.text,4,2))+Ext
    else
    NomFicExport.text := Copy(NomFicExport.text,1,pos(Ext1,NomFicExport.text)-1)+Ext;
  if NomFicIntro.text = '' then
    NomFicIntro.text := 'C:\INSTAT\PGII'+trim(Copy(DateDebInt.text,4,2))+Ext
    else
    NomFicIntro.text := Copy(NomFicIntro.text,1,pos(Ext1,NomFicIntro.text)-1)+Ext;
  if FRMT_EXC.Checked then
    Ext1 := '.xls'
    else
    Ext1 := '.sdf';
end;

//========================================
// Récapitulatif des choix à l'écran
//========================================
procedure TFAssistDEBC.ListeRecap;
Var st_chaine : string;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
NbItems := 2;
if CB_CLIENTS.checked then
   BEGIN
   st_chaine := '* Extraction des clients modifiés depuis le '+CDATECLI.text;
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  dans le fichier '+NomFicCli.text;
   ListRecap.Items.Add (st_chaine);
   NbItems := NbItems + 2;
   end;

if CB_ARTICLES.Checked then
   BEGIN
   st_chaine := '* Extraction des articles créés depuis le '+CDATEART.text;
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  dans le fichier '+NomFicArt.text;
   ListRecap.Items.Add (st_chaine);
   NbItems := NbItems + 2;
   end;

if CB_EXPORT.Checked then
   BEGIN
   st_chaine := '* Extraction des lignes pour la D.E.B à l''exportation du mois '+
       trim(Copy(DateDebExp.text,4,2));
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  dans le fichier '+NomFicExport.text;
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  L''unité de mesure correspondant au Kilo est "'+KiloE.text+'"';
   ListRecap.Items.Add (st_chaine);
   NbItems := NbItems + 3;
   if MontantLibreE.text <> '<<Aucun>>' then
      begin
      st_chaine := '  Vous Utilisez le champ "'+MontantLibreE.text+'" pour les unités suppl.';
      ListRecap.Items.Add (st_chaine);
      NbItems := NbItems + 1;
      end;
   end;

if CB_INTRO.Checked then
   BEGIN
   st_chaine := '* Extraction des lignes pour la D.E.B à l''introduction du mois '+
       trim(Copy(DateDebInt.text,4,2));
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  dans le fichier '+NomFicIntro.text;
   ListRecap.Items.Add (st_chaine);
   st_chaine := '  L''unité de mesure correspondant au Kilo est "'+KiloI.text+'"';
   ListRecap.Items.Add (st_chaine);
   NbItems := NbItems + 3;
   if MontantLibreI.text <> '<<Aucun>>' then
      begin
      st_chaine := '  Vous Utilisez le champ "'+MontantLibreI.text+'" pour les unités suppl.';
      ListRecap.Items.Add (st_chaine);
      NbItems := NbItems + 1;
     end;
   end;
ListRecap.Items.Add ('');
NbItems := NbItems + 1;
END;

//===========================================================
// sur le bouton fin : contrôles et lancement du traitement
//===========================================================

procedure TFAssistDEBC.bFinClick(Sender: TObject);
var DD : TDateTime;
    DF : TDatetime;
    Ok : boolean;
begin
  inherited;
  if CB_EXPORT.Checked then
     begin
     DD := StrToDate(DateDebExp.Text);
     DF := StrToDate(DateFinExp.Text);
     if DF < DD then
        begin
        MsgBox.execute(0,caption,'');
        exit;
        end;
     if trim(Copy(DateDebExp.text,4,7)) <> trim(Copy(DateFinExp.Text,4,7)) then
        begin
        MsgBox.execute(1,caption,'');
        exit;
        end;
     if KiloE.text = '' then
        begin
        MsgBox.execute(2,caption,'');
        exit;
        end;
     end;
  if CB_INTRO.Checked then
     begin
     DD := StrToDate(DateDebInt.Text);
     DF := StrToDate(DateFinInt.Text);
     if DF < DD then
        begin
        MsgBox.execute(3,caption,'');
        exit;
        end;
     if trim(Copy(DateDebInt.text,4,7)) <> trim(Copy(DateFinInt.Text,4,7)) then
        begin
        MsgBox.execute(4,caption,'');
        exit;
        end;
     if KiloI.text = '' then
        begin
        MsgBox.execute(2,caption,'');
        exit;
        end;
     end;
  Ok := true;
  AuMoinsUn := false;
  ListeRecap;
  ListRecap.Items.Add('*** Compte rendu du traitement ***');
  ListRecap.Items.Add('');

  // On va ensuite demander à toutes les nouvelles lignes que la ligne
  // *** Compte rendu du traitement *** soit placée au top : cela ne se fera
  // que quand cela sera possible mais, pour une bonne approche de la chose,
  // les lignes vont scroller jusqu'à ce point
  ListRecap.TopIndex := NbItems;

  if CB_CLIENTS.Checked then
    Ok := TraitementClients;
  if (Ok) and (CB_ARTICLES.Checked) then
    Ok := TraitementArticles;
  if (Ok) and (CB_EXPORT.Checked) then
    Ok := TraitementDEB('E');
  if (Ok) and (CB_INTRO.Checked) then
    Ok := TraitementDEB('I');
  if Ok then MsgBox.execute(13,caption,'');
  if AuMoinsUn then EnregistreEvenement;
  bFermer.Visible := True;
  bFin.Visible := False;

end;

//==================================================
// Génération du fichier clients
//===================================================
Function TFAssistDEBC.TraitementClients : boolean;
var FicOut : TextFile;
    TobClients, TobExcel, TobTmp : TOB;
    QQ : TQuery;
    DD : TDateTime;
    nbenr : Integer;
    Libelle : string;
    Nif : string;
    NomFic : string;
    TobLig : TOB;
    ind : Integer;
    rep : integer;
BEGIN
  Result := true;
  ListRecap.Items.Add('* Extraction des clients');
  ListRecap.TopIndex := NbItems;
  SourisSablier;
  DD := StrToDate(CDATECLI.Text);
  TobClients := TOB.Create('',Nil,-1);
  QQ := OpenSql('SELECT T_NIF, T_LIBELLE from TIERS WHERE T_NATUREAUXI="CLI" AND T_DATEMODIF >= "'
        + UsdateTime(DD) + '" AND T_NIF <> ""', true);
  if not QQ.Eof then
    TobClients.LoadDetailDB('TIERS','','',QQ,false)
    else
    FreeAndNil(TobClients);
  Ferme(QQ);
  if TobClients = Nil then
  begin
    ListRecap.Items.Add('  Aucun client sélectionné...');
    ListRecap.TopIndex := NbItems;
    SourisNormale;
    exit;
  end;
  if Excel then
  begin
    NomFic := NomFicCli.text;
    TobExcel := TOB.Create('',nil,-1);
  end else
  begin
    // ouverture fichier OUT
    AssignFile(FicOut, NomFicCli.text);
    {$I-} Reset(FicOut); {$I+}
    if IoResult = 0 then
    begin
      CloseFile(FicOut);
      rep := MsgBox.execute(5,caption,'');
      if rep <> MrYes then
      begin
        SourisNormale;
        TobClients.free;
        if rep = MrCancel then
        begin
          ListRecap.Items.Add('  Traitement interrompu par l''utilisateur...');
          result := false;
        end else
        begin
          ListRecap.Items.Add('  Abandonnée...');
        end;
        ListRecap.TopIndex := NbItems;
        Exit;
      end;
    end;
    AssignFile(FicOut,NomFicCli.text);
    {$I-} Rewrite(FicOut); {$I+}
    if IoResult <> 0 then
    begin
      ListRecap.Items.Add('  Impossible de créer le fichier '+NomFicCli.text);
      ListRecap.TopIndex := NbItems;
      SourisNormale;
      TobClients.free;
      Exit;
    end;
  end;
  nbenr := TobClients.Detail.Count;
  if nbenr <> 0 then InitMoveProgressForm (Self,'Export Fichier Clients','Préparation des données ',nbenr,false,True) ;
  nbenr := 0;
  for ind := 0 to TobClients.Detail.count - 1 do
  begin
    TobLig := TobClients.Detail[ind];
    nbenr := nbenr+1;
    Libelle := TobLig.GetValue('T_LIBELLE');
    MoveCurProgressForm ('Ecriture du client : '+Libelle);
    Nif := TobLig.GetValue('T_NIF');
    FormatChaine(Nif,20);
    FormatChaine(Libelle,35);
    if Excel then
    begin
      TobTmp := TOB.Create('',TobExcel,-1);
      TobTmp.AddChampSupValeur('NIF',Nif,false);
      TobTmp.AddChampSupValeur('LIBELLE',Libelle,false);
    end else
    begin
      Writeln (FicOut,Nif+Libelle);
    end;
    AuMoinsUn := true;
  end;
  if nbenr = 1 then
    ListRecap.Items.Add('  1 client extrait')
    else
    ListRecap.Items.Add('  '+IntToStr(nbenr)+' clients extraits');
  ListRecap.TopIndex := NbItems;
  if Excel then
  begin
//    TobExcel.SaveToExcelFile(NomFic);
    ExportTobToExcel(TobExcel,NomFic,False);
    FreeAndNil(TobExcel);
  end else
  begin
    CloseFile(FicOut);
  end;
  FiniMoveProgressForm ;
  TobClients.free;
  SourisNormale;
end;

//==================================================
// Génération du fichier Articles
//===================================================
function TFAssistDEBC.TraitementArticles : boolean;
var FicOut : TextFile;
    TobArticles, TobExcel, TobTmp : TOB;
    QQ : TQuery;
    DD : TDateTime;
    nbenr : Integer;
    Libelle : string;
    Article : string;
    Nc8, Ngp9 : string;
    NomFic : string;
    TobLig : TOB;
    ind : Integer;
    rep : integer;
BEGIN
  result := true;
  ListRecap.Items.Add('* Extraction des articles');
  ListRecap.TopIndex := NbItems;
  SourisSablier;
  DD := StrToDate(CDATEART.Text);
  TobArticles := TOB.Create('',Nil,-1);
  QQ := OpenSql('SELECT GA_CODEARTICLE, GA_LIBELLE, GA_CODEDOUANIER from ARTICLE WHERE (GA_STATUTART="GEN" OR GA_STATUTART="UNI") AND GA_DATECREATION >= "'
     +UsdateTime(DD)+'" AND GA_CODEDOUANIER <> "" AND ((GA_TYPEARTICLE= "MAR") OR (GA_TYPEARTICLE="NOM" AND GA_TYPENOMENC="ASS"))'
     ,true);
  if not QQ.Eof then
    TobArticles.LoadDetailDB('ARTICLE','','',QQ,false)
    else
    FreeAndNil(TobArticles);
  Ferme(QQ);
  if TobArticles = Nil then
  begin
    ListRecap.Items.Add('  Aucun article sélectionné');
    ListRecap.TopIndex := NbItems;
    SourisNormale;
    exit;
  end;
  if Excel then
  begin
    NomFic := NomFicArt.text;
    TobExcel := TOB.Create('',nil,-1);
  end else
  begin
    AssignFile(FicOut,NomFicArt.text);
    {$I-} Reset(FicOut); {$I+}
    if IoResult = 0 then
    begin
      CloseFile(FicOut);
      rep := MsgBox.execute(9,caption,'');
      if rep <> MrYes then
      begin
        SourisNormale;
        TobArticles.free;
        if rep = MrCancel then
        begin
          ListRecap.Items.Add('  Traitement interrompu par l''utilisateur...');
          result := false;
        end else
        begin
          ListRecap.Items.Add('  Abandonnée...');
        end;
        ListRecap.TopIndex := NbItems;
        Exit;
      end;
    end;
    AssignFile(FicOut,NomFicArt.text);
    {$I-} Rewrite(FicOut); {$I+}
    if IoResult <> 0 then
    begin
      ListRecap.Items.Add('  Impossible de créer le fichier '+ NomFicArt.text);
      ListRecap.TopIndex := NbItems;
      SourisNormale;
      TobArticles.free;
      Exit;
    end;
  end;
  nbenr := TobArticles.Detail.Count;
  if nbenr <> 0 then
    InitMoveProgressForm (Self,'Export Fichier Articles','Préparation des données ',nbenr,false,True) ;
  nbenr := 0;
  for ind := 0 to TobArticles.Detail.count - 1 do
  begin
    TobLig := TobArticles.Detail[ind];
    nbenr := nbenr+1;
    MoveCurProgressForm ('Ecriture de l''article : '+TobLig.GetValue('GA_CODEARTICLE')) ;
    Article := TobLig.GetValue('GA_CODEARTICLE'); FormatChaine(Article,25);
    Nc8 := TobLig.GetValue('GA_CODEDOUANIER');    FormatChaine(Nc8,8);
    Ngp9 := Copy(TobLig.GetValue('GA_CODEDOUANIER'),9,1);
    Libelle := TobLig.GetValue('GA_LIBELLE');     FormatChaine(Libelle,140);
    if Excel then
    begin
      TobTmp := TOB.Create('',TobExcel,-1);
      TobTmp.AddChampSupValeur('ARTICLE',Article,false);
      TobTmp.AddChampSupValeur('NC8',Nc8,false);
      TobTmp.AddChampSupValeur('NGP9',Ngp9,false);
      TobTmp.AddChampSupValeur('LIBELLE',Libelle,false);
    end else
    begin
      Writeln (FicOut,Article+Nc8+Ngp9+Libelle);
    end;
    AuMoinsUn := true;
  end;
  if nbenr = 1 then
    ListRecap.Items.Add('  1 article extrait')
    else
    ListRecap.Items.Add('  '+IntToStr(nbenr)+' articles extraits');
  ListRecap.TopIndex := NbItems;
  if Excel then
  begin
//    TobExcel.SaveToExcelFile(NomFic);
    ExportTobToExcel(TobExcel,NomFic,False);
    FreeAndNil(TobExcel);
  end else
  begin
    CloseFile(FicOut);
  end;
  FiniMoveProgressForm ;
  TobArticles.free;
  SourisNormale;
end;

//================================================
// Génération des lignes de DEB
//================================================
Function TFAssistDEBC.TraitementDEB(TypeDEB : string) : boolean;
var TOBGlobal, TobExcel, TobTmp : TOB;
    FicOut : TextFile;
    WhereLigne : string;
    NomFicDeb : string;
    TitreProgress : string;
    Numess : integer;
    TobPiece : TOB;
    TobLigne : TOB;
    ind1 : Integer;
    ind2 : Integer;
    QQ : TQuery;
    ChercheMea : string;
    NuLibre : integer;
    NomLibre : string;
    rep : integer;
    nblig : integer;
    // entete
    Transaca : string;
    Transacb : string;
    Regime : string;
    PaysCee : string;
    IdentCee : string;
    Devise : string;
    Monnaie : string;
    Nudoc : string;
    // detail
    CodeDouane : string;
    CodeDouaneSupp : string; //JT
    PaysOri : string;
    Depart : string;
    Poids : string;
    Montant,MontantStat : string;
    MontantDev,MontantStatDev : string;
    UniteSup : string;
    Transport : string;
    Incoterm : string;
    LieuDispo : string;
    pp : double;
    mt : double;
    mts : double; //JT
    md : double;
    mtsd : double; //JT
    us : double;
    qte : double;
    qot : double;
    qok : double;
    ml : double;
begin
  Result := true;
  if TypeDEB = 'E' then
    ListRecap.Items.Add('* Extraction des factures clients')
    else
    ListRecap.Items.Add('* Extraction des factures fournisseurs');
  ListRecap.TopIndex := NbItems;
  SourisSablier;
  TOBGlobal := Nil;
  ChargerLesPieces(TypeDEB,TOBGlobal,WhereLigne);
  if (TOBGlobal = Nil) or (TOBGlobal.Detail.count = 0) then
  begin
    ListRecap.Items.Add('  Aucune facture sélectionnée');
    ListRecap.TopIndex := NbItems;
    SourisNormale;
    exit;
  end;
  ChargerLesLignes(TypeDEB,TOBGlobal,WhereLigne,nblig);
  if (TOBGlobal = Nil) or (TOBGlobal.Detail.count = 0) then
  begin
    ListRecap.Items.Add('  Aucune ligne sélectionnée');
    ListRecap.TopIndex := NbItems;
    SourisNormale;
    exit;
  end;

  // ouverture fichier OUT
  If Typedeb = 'E' then
  begin
    NomFicDeb := NomFicExport.text;
    Numess := 14;
    end else
  begin
    NomFicDeb := NomFicIntro.text;
    Numess := 15;
  end;

  // Préparation de la Tob vers Excel ou du fichier d'export
  if Excel then
  begin
    TobExcel := TOB.Create('',nil,-1);
  end else
  begin
    AssignFile(FicOut,NomFicDeb);
    {$I-} Reset(FicOut); {$I+}
    if IoResult = 0 then
    begin
      CloseFile(FicOut);
      rep := MsgBox.execute(numess,caption,'');
      if rep <> MrYes then
      begin
        SourisNormale;
        TOBGlobal.free;
        if rep = MrCancel then
        begin
          ListRecap.Items.Add('  Traitement interrompu par l''utilisateur...');
          result := false;
        end else
        begin
          ListRecap.Items.Add('  Abandonnée...');
        end;
        ListRecap.TopIndex := NbItems;
        Exit;
      end;
    end;
    AssignFile(FicOut,NomFicDeb);
   {$I-} Rewrite(FicOut); {$I+}
    if IoResult <> 0 then
    begin
      ListRecap.Items.Add('  Impossible de créer le fichier '+NomFicDeb);
      ListRecap.TopIndex := NbItems;
      SourisNormale;
      TOBGlobal.Free;
      Exit;
    end;
  end;

  // accès à la table mea pour quotité du kg
  if TypeDEB = 'E' then
     ChercheMea := 'SELECT GME_QUOTITE FROM MEA WHERE GME_QUALIFMESURE="POI" AND GME_MESURE="'+KiloE.Value+'"'
  else
     ChercheMea := 'SELECT GME_QUOTITE FROM MEA WHERE GME_QUALIFMESURE="POI" AND GME_MESURE="'+KiloI.Value+'"';
  QQ := OpenSql(ChercheMea,true);
  if not QQ.Eof then
    qok := QQ.Findfield('GME_QUOTITE').AsFloat
    else
    qok := 1;
  Ferme(QQ);

  // traitement de la tob constituée
  If TypeDEB = 'E' then
    TitreProgress := 'Déclaration à l''exportation'
    else
    TitreProgress := 'Déclaration à l''Introduction';
  InitMoveProgressForm(Self,TitreProgress,'Ecriture des données ',TOBGlobal.Detail.count,false,True) ;
  for ind1 := 0 to TOBGlobal.Detail.count - 1 do
  begin
    MoveCur(False);
    TobPiece  := TOBGlobal.detail[ind1];
    TransacA  := TobPiece.GetValue('TransacA');        FormatChaine(TransacA,1);
    TransacB  := TobPiece.GetValue('TransacB');        FormatChaine(TransacB,1);
    Regime    := TobPiece.GetValue('Regime');          FormatChaine(Regime,5);
    PaysCee   := TobPiece.GetValue('PY_CODEBANCAIRE'); FormatChaine(PaysCee,3);
    IdentCee  := TobPiece.GetValue('T_NIF');           FormatChaine(IdentCee,20);
    Transport := TobPiece.GetValue('YTC_MODEEXP');     FormatChaine(Transport,1); //JT
    Incoterm  := TobPiece.GetValue('YTC_INCOTERM');    FormatChaine(Incoterm,3); //JT
    LieuDispo := TobPiece.GetValue('YTC_LIEUDISPO');   FormatChaine(LieuDispo,1); //JT
    Nudoc     := TobPiece.GetValue('GP_SOUCHE')+IntToStr(TobPiece.GetValue('GP_NUMERO')); FormatChaine(Nudoc,14);
    if TobPiece.GetValue('GP_DEVISE') <> GetParamSoc('SO_DEVISEPRINC') then
    begin
      Devise := TobPiece.GetValue('D_SYMBOLE');
    end else
    begin
      if TobPiece.GetValue('GP_SAISIECONTRE') <> 'X' then
      begin
        Devise := TobPiece.GetValue('D_SYMBOLE');
      end else
      begin
        if VH^.TenueEuro then
          Devise := 'FRF'
          else
          Devise := 'EUR';
      end;
    end;
    FormatChaine(Devise,3);
    if VH^.TenueEuro then
      Monnaie := 'EUR'
      else
      Monnaie := 'FRF';
    FormatChaine(Monnaie,3);
    for ind2 := 0 to TobPiece.Detail.count - 1 do
    begin
      TobLigne := TobPiece.Detail[ind2];
      CodeDouane := TobLigne.GetValue('GL_CODEDOUANE'); FormatChaine(CodeDouane,8);
      CodeDouaneSupp := Copy(TobLigne.GetValue('GL_CODEDOUANE'),9,1); // JT DEB
      if TypeDEB = 'E' then
        PaysOri := ''
        else
        PaysOri := TobPiece.GetValue('PY_CODEBANCAIRE');
      FormatChaine(PaysOri,3);
      // ATTENTION, Exception pour MONACO (CP = 98000 mais utiliser 99)
      if UpperCase(TobLigne.GetValue('GDE_VILLE')) = 'MONACO' then
        Depart := '99'
        else
        Depart := TobLigne.GetValue('GDE_CODEPOSTAL');
      FormatChaine(Depart,2);
      if TobLigne.GetValue ('GME_QUOTITE') <> NULL then
        qot := TobLigne.GetValue('GME_QUOTITE')
        else
        qot := 1;
      pp := TobLigne.GetValue('GL_TOTALPOIDSDOUA');
      pp := (pp * qot)/qok;                       Poids := Format('%15f',[pp]);
      mt := TobLigne.GetValue('GL_TOTALHT');      Montant := Format('%13.2f',[mt]);
      mts := TobLigne.GetValue('GL_TOTALHT');     MontantStat := Format('%13.2f',[mts]); //JT
      md := TobLigne.GetValue('GL_TOTALHTDEV');   MontantDev := Format('%13.2f',[md]);
      mtsd := TobLigne.GetValue('GL_TOTALHTDEV'); MontantStatDev := Format('%13.2f',[mtsd]); //JT
      Nulibre := 0;
      if TypeDeb = 'E' then
      begin
        if MontantLibreE.Text <> '<<Aucun>>' then
           NuLibre := strtoint (Copy (MontantLibreE.Text, length (MontantLibreE.Text), 1));
      end else
      begin
        if MontantLibreI.Text <> '<<Aucun>>' then
           NuLibre := strtoint (Copy (MontantLibreI.Text, length (MontantLibreI.Text), 1));
      end;
      if NuLibre <> 0 then
      begin
        NomLibre := 'GA_VALLIBRE'+IntToStr(NuLibre);
        ml := TobLigne.GetValue(NomLibre);
        qte := TobLigne.GetValue('GL_QTEFACT');
        us := ml * qte;
        UniteSup := Format('%13.2f',[us]);
      end else
      begin
        UniteSup := '';
        FormatChaine(UniteSup,13);
      end;
      if Excel then
      begin
        TobTmp := TOB.Create('',TobExcel,-1);
        TobTmp.AddChampSupValeur('CODEDOUANE',CodeDouane,false);
        TobTmp.AddChampSupValeur('CODEDOUANESUPP',CodeDouaneSupp,false);
        TobTmp.AddChampSupValeur('TRANSACA',TransacA,false);
        TobTmp.AddChampSupValeur('TRANSACB',TransacB,false);
        TobTmp.AddChampSupValeur('REGIME',Regime,false);
        TobTmp.AddChampSupValeur('_VIDE1','',false);
        TobTmp.AddChampSupValeur('PAYSCEE',PaysCee,false);
        TobTmp.AddChampSupValeur('PAYSORI',PaysOri,false);
        TobTmp.AddChampSupValeur('DEPART',Depart,false);
        TobTmp.AddChampSupValeur('_VIDE2','',false);
        TobTmp.AddChampSupValeur('POIDS',Poids,false);
        TobTmp.AddChampSupValeur('UNITESUP',UniteSup,false);
        TobTmp.AddChampSupValeur('IDENTCEE',IdentCee,false);
        TobTmp.AddChampSupValeur('TRANSPORT',Transport,false);
        TobTmp.AddChampSupValeur('MONTANT',Montant,false);
        TobTmp.AddChampSupValeur('MONTANTSTAT',MontantStat,false);
        TobTmp.AddChampSupValeur('_VIDE3','',false);
        TobTmp.AddChampSupValeur('NUDOC',Nudoc,false);
        TobTmp.AddChampSupValeur('INCOTERM',Incoterm,false);
        TobTmp.AddChampSupValeur('_VIDE4','',false);
        TobTmp.AddChampSupValeur('_VIDE5','',false);
        TobTmp.AddChampSupValeur('LIEUDISPO',LieuDispo,false);
        TobTmp.AddChampSupValeur('MONTANTDEV',MontantDev,false);
        TobTmp.AddChampSupValeur('MONNAIE',Monnaie,false);
        TobTmp.AddChampSupValeur('MONTANTSTATDEV',MontantStatDev,false);
      end else
      begin
        Writeln (FicOut,CodeDouane+CodeDouaneSupp+TransacA+TransacB+Regime+PaysCee+PaysOri+Depart+Poids+UniteSup+IdentCee+Transport+Montant+MontantStat+Nudoc+Incoterm+LieuDispo+MontantDev+Monnaie+MontantStatDev)
      end;
      AuMoinsUn := true;
    end;
  end;
  if nblig = 1 then
    ListRecap.Items.Add('  1 ligne extraite')
    else
    ListRecap.Items.Add('  '+IntToStr(nblig)+' lignes extraites');
  ListRecap.TopIndex := NbItems;
  if Excel then
  begin
//    TobExcel.SaveToExcelFile(NomFicDeb);
    ExportTobToExcel(TobExcel,NomFicDeb,False);
    FreeAndNil(TobExcel);
  end else
  begin
    CloseFile(FicOut);
  end;
  FiniMoveProgressForm ;
  TobGlobal.free;
  SourisNormale;
end;

//=================================================
// première étape : charger les pièces concernées
//=================================================
Procedure TFAssistDEBC.ChargerLesPieces(TypeDEB : string; var TOBGlobal : TOB; var WhereLigne : string);
var LeBigSQL : string;
    WhereNat : string;
    TOBNat : TOB;
    DD : TDateTime;
    DF : TDateTime;
    QQ : TQuery;
    ind : Integer;
    TOBPiece : TOB;
    prem : boolean;
    TN : TOB;
    TitreProgres : string;
    nbenr : Integer;
BEGIN

// construction du SQL
if TypeDEB = 'E' then
   begin
   DD := StrToDate(DateDebExp.Text);
   DF := StrToDate(DateFinExp.Text);
   end else begin
   DD := StrToDate(DateDebInt.Text);
   DF := StrToDate(DateFinInt.Text);
   end;
TOBNat := Nil;
WhereParPiece(TypeDEB, TOBNat, WhereNat);
If TOBNat = Nil  then exit;
LeBigSQL := 'SELECT GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG, GP_CONTREMARQUE, GP_SAISIECONTRE, GP_DEVISE, ';
LeBigSQL := LeBigSQL+'YTC_MODEEXP, YTC_INCOTERM, YTC_LIEUDISPO, '; //JT
LeBigSQL := LeBigSQL+'PY_CODEBANCAIRE, T_NIF, D_SYMBOLE FROM PIECE LEFT JOIN TIERS ON T_TIERS=GP_TIERS ';
if TypeDEB = 'E' then
   begin
   if GetParamSoc('SO_GCPIECEADRESSE') then
      begin
      LeBigSQL := LeBigSQL+'INNER JOIN PIECEADRESSE ON GPA_NATUREPIECEG=GP_NATUREPIECEG AND ';
      LeBigSQL := LeBigSQL+'GPA_SOUCHE=GP_SOUCHE AND GPA_NUMERO=GP_NUMERO AND GPA_INDICEG=GP_INDICEG AND ';
      LeBigSQL := LeBigSQL+'GPA_NUMLIGNE=0 AND GPA_TYPEPIECEADR="001" ';
      LeBigSQL := LeBigSQL+'INNER JOIN PAYS ON PY_PAYS=GPA_PAYS ';
      end else
      begin
      LeBigSQL := LeBigSQL+'INNER JOIN ADRESSES ON ADR_TYPEADRESSE="PIE" AND ADR_NUMEROADRESSE=GP_NUMADRESSELIVR ';
      LeBigSQL := LeBigSQL+'INNER JOIN PAYS ON PY_PAYS=ADR_PAYS ';
      end;
   end else begin
   LeBigSQL := LeBigSQL+'INNER JOIN PAYS ON PY_PAYS=T_PAYS ';
   end;
LeBigSQL := LeBigSQL+'LEFT JOIN DEVISE ON D_DEVISE=GP_DEVISE ';
LeBigSQL := LeBigSQL+'LEFT JOIN TIERSCOMPL ON YTC_TIERS=GP_TIERS '; //JT
LeBigSQL := LeBigSQL+'WHERE '+WhereNat+' AND GP_DATEPIECE >= "'+UsdateTime(DD);
LeBigSQL := LeBigSQL+'" AND GP_DATEPIECE <= "'+UsdateTime(DF);
LeBigSQL := LeBigSQL+'" AND PY_MEMBRECEE="X" AND PY_CODEBANCAIRE <> "001"';
TOBGlobal := TOB.Create('',Nil,-1);
QQ := OpenSql(LeBigSQL,true);
if not QQ.Eof then TOBGlobal.LoadDetailDB('','','',QQ,false)
              else begin TOBGlobal.Free; TOBGlobal:=Nil; end;
Ferme(QQ);
If TOBGlobal = Nil then
   begin
   TOBNat.free;
   exit;
   end;

// On retraite la TOB chargée pour obtenir une clause where sur les lignes
// en même temps, on complète les renseignements à transmettre à partir de TOBNat
nbenr := TobGlobal.detail.count;
If TypeDEB = 'E' then TitreProgres := 'Déclaration à l''exportation'
else TitreProgres := 'Déclaration à l''Introduction';
if nbenr <> 0 then InitMoveProgressForm (Self,TitreProgres,'Préparation des données ',nbenr,false,True) ;
nbenr := 0;
prem := true;
for ind := 0 to TOBGlobal.Detail.count - 1 do
    begin
    nbenr := nbenr+1;
    MoveCurProgressForm ('chargement des pièces  ');
    TOBPiece := TOBGlobal.Detail[ind];
    TOBPiece.AddChampSup('Regime',False);
    TOBPiece.AddChampSup('TransacA',False);
    TOBPiece.AddChampSup('TransacB',False);
    TN := TOBNat.FindFirst(['GPP_NATUREPIECEG'],[TOBPiece.GetValue('GP_NATUREPIECEG')],true);
    TOBPiece.PutValue('Regime',TN.GetValue('Regime'));
    TOBPiece.PutValue('TransacA',TN.GetValue('TransacA'));
    TOBPiece.PutValue('TransacB',TN.GetValue('TransacB'));
    if prem then WhereLigne := '(('
    else WhereLigne := WhereLigne+' OR (';
    prem := false;
    WhereLigne := WhereLigne+'GL_NATUREPIECEG="'+TOBPiece.GetValue('GP_NATUREPIECEG')+
      '" AND GL_SOUCHE="'+TOBPiece.GetValue('GP_SOUCHE')+'" AND GL_NUMERO='+
      IntToStr(TOBPiece.GetValue('GP_NUMERO'))+' AND GL_INDICEG='+
      IntToStr(TOBPiece.GetValue('GP_INDICEG'))+')';
    end;
if not prem then WhereLigne := WhereLigne+')';
FiniMoveProgressForm ;
TOBNat.free;
end;

//===============================================
// Chargement des lignes sélectionnées
//===============================================
Procedure TFAssistDEBC.ChargerLesLignes(TypeDEB :string; var TOBGlobal : TOB; WhereLigne : string; var nblig : integer);
var TOBTmp : TOB;
    LeGrosSQL : string;
    NuLibre : integer;
    WhereNom : string;
    QQ : TQuery;
    TT : TOB;
    TobPiece : TOB;
    TobLigne : TOB;
    ind : Integer;
    Nomlue : boolean;
    TOBNomenc : TOB;
    TOBCompos : TOB;
    TitreProgres : string;
    nbenr : Integer;
    Cle1 : string;
    Cle2 : string;
BEGIN
// construction du sql
LeGrosSQL := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_CODEDOUANE, GL_QUALIFPOIDS';
LeGrosSQL := LeGrosSQL+' GL_QTEFACT, GL_NUMLIGNE, GL_ARTICLE, GL_NUMORDRE,';
LeGrosSQL := LeGrosSQL+' GL_REFARTTIERS, GL_TOTALPOIDSDOUA, GL_TOTALHT, GL_TOTALHTDEV,';
LeGrosSQL := LeGrosSQL+' GL_CODEARTICLE, GL_TYPEARTICLE, GL_REFARTBARRE, GDE_CODEPOSTAL, GDE_VILLE, GME_QUOTITE';
if TypeDEB = 'I' then LeGrosSQL := LeGrosSQL+', PY_CODEBANCAIRE';
NuLibre := 0;
if TypeDeb = 'E' then
   begin
   if MontantLibreE.Text <> '<<Aucun>>' then
         NuLibre := strtoint (Copy (MontantLibreE.Text, length (MontantLibreE.Text), 1));
   end else begin
   if MontantLibreI.Text <> '<<Aucun>>' then
      NuLibre := strtoint (Copy (MontantLibreI.Text, length (MontantLibreI.Text), 1));
   end;
if NuLibre <> 0 then
   begin
   LeGrosSQL := LeGrosSQL+', GA_VALLIBRE'+IntToStr(NuLibre);
   end;
LeGrosSQL := LeGrosSQL+' FROM LIGNE ';
if NuLibre <> 0 then
   LeGrosSQL := LeGrosSQL+'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE ';
LeGrosSQL := LeGrosSQL+' LEFT JOIN DEPOTS ON GDE_DEPOT=GL_DEPOT';
LeGrosSQL := LeGrosSQL+' LEFT JOIN MEA ON GME_QUALIFMESURE="POI" AND GME_MESURE=GL_QUALIFPOIDS';
If TypeDEB = 'I' then
   LeGrosSQL := LeGrosSQL+' LEFT JOIN PAYS ON PY_PAYS=GL_PAYSORIGINE';
LeGrosSQL := LeGrosSQL+' WHERE '+WhereLigne+' AND GL_TYPELIGNE="ART"';
LeGrosSQL := LeGrosSQL+' AND (GL_TYPEARTICLE="MAR" OR GL_TYPEARTICLE="NOM" OR GL_TYPEREF="CAT")';
LeGrosSQL := LeGrosSQL+' AND GL_TOTALHTDEV <> 0 ORDER BY GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG';
TOBTmp := TOB.Create('',Nil,-1);
QQ := OpenSql(LeGrosSQL,true);
if not QQ.Eof then TOBTmp.LoadDetailDB('','','',QQ,false)
else TOBTmp := Nil;
Ferme(QQ);
If TOBTmp = Nil then
   begin
   TOBGlobal.free;
   TOBGlobal := Nil;
   exit;
   end;

// On affecte les lignes à leur maman et on en profite pour éliminer
// Les lignes <> MAR OU NOM mais pas en contremarque et pour éclater
// les nomenclatures jusqu'au premier NC8

nbenr := TobTmp.Detail.count;
If TypeDEB = 'E' then TitreProgres := 'Déclaration à l''exportation'
else TitreProgres := 'Déclaration à l''Introduction';
if nbenr <> 0 then InitMoveProgressForm (Self,TitreProgres,'Préparation des données ',nbenr,false,True) ;
nbenr := 0;
Nomlue := false;
TOBNomenc := Nil;
nblig := 0;
Cle1 := '';
Cle2 := '';
for ind := 0 to TOBTmp.Detail.count - 1 do
    begin
    nbenr := nbenr+1;
    MoveCurProgressForm ('chargement des lignes de pièce ');
    TT := TOBTmp.Detail[ind];
    Cle1 := TT.GetValue('GL_NATUREPIECEG')+TT.GetValue('GL_SOUCHE')+IntToStr(TT.GetValue('GL_NUMERO'))
            +IntToStr(TT.GetValue('GL_INDICEG'));
    If Cle1 <> Cle2 then
       begin
       If TOBNomenc <> Nil then TOBNomenc.free;
       NomLue := false;
       Cle2 := Cle1;
       TobPiece := TOBGlobal.FindFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
             [TT.GetValue('GL_NATUREPIECEG'),TT.GetValue('GL_SOUCHE'),TT.GetValue('GL_NUMERO')
             ,TT.GetValue('GL_INDICEG')],False);
       end;
    if (TT.GetValue('GL_TYPEARTICLE') = 'MAR') OR (TT.GetValue('GL_TYPEARTICLE') = 'NOM')
       OR (TobPiece.GetValue('GP_CONTREMARQUE') = 'X') then
       begin
       if (TT.GetValue('GL_TYPEARTICLE') = 'NOM') AND (TT.GetValue('GL_CODEDOUANE') = '') then
          begin
          // pour ne pas faire la requête s'il n'y a pas de nomenclature à traiter
          // dans la pièce, on ne charge les nomenclatures qu'au premier cas rencontré
          if not NomLue then
             begin
             TOBNomenc := Nil;
             WhereNom := '(GLN_NATUREPIECEG="'+TOBPiece.GetValue('GP_NATUREPIECEG')+
                   '" AND GLN_SOUCHE="'+TOBPiece.GetValue('GP_SOUCHE')+'" AND GLN_NUMERO='+
                   IntToStr(TOBPiece.GetValue('GP_NUMERO'))+' AND GLN_INDICEG='+
                   IntToStr(TOBPiece.GetValue('GP_INDICEG'))+')';
             ChargerLesNoms(TypeDEB,WhereNom,TOBNomenc);
             Nomlue := true;
             end;
          // traiter la nomenclature
          TOBCompos := Nil;
          Eclater(TOBCompos,TOBNomenc,TT.GetValue('GL_NUMLIGNE'),0,TT.GetValue('GL_ARTICLE'),0,
               TT.GetValue('GL_QTEFACT'));
          If TOBCompos = Nil then
             begin
             TobLigne := TOB.Create('',TobPiece,-1);
             TobLigne.Dupliquer(TT,false,True);
             nblig := nblig + 1;
             end else begin
             TobEclatToTobLigne(TypeDEB, TobPiece,TT,TOBCompos,NbLig);
             TobCompos.free;
             end;
          end else begin
          TobLigne := TOB.Create('',TobPiece,-1);
          TobLigne.Dupliquer(TT,false,True);
          nblig := nblig + 1;
          end;
       end;
    end;
FiniMoveProgressForm ;
TOBNomenc.free;
TOBTmp.free;
end;

//=============================================================
// chargement des lignes de nomenclature (pour toute la piece)
//=============================================================
Procedure TFAssistDEBC.ChargerLesNoms(TypeDEB : string; WhereNom : string; var TOBNomenc : TOB);
var LeNomSQL : string;
    NuLibre : integer;
    QQ : TQuery;
begin
// construction du sql
LeNomSQL := 'SELECT GLN_NATUREPIECEG, GLN_SOUCHE, GLN_NUMERO, GLN_INDICEG, GLN_NUMLIGNE, GLN_ARTICLE,';
LeNomSQL := LeNomSQL+' GLN_COMPOSE, GLN_NIVEAU, GLN_CODEARTICLE, GLN_NUMORDRE, GLN_ORDRECOMPO,';
LeNomSQL := LeNomSQL+' GLN_QTE, GA_CODEDOUANIER, GA_POIDSDOUA, GA_PVHT, GA_TYPEARTICLE,';
LeNomSQL := LeNomSQL+' GA_CODEBARRE, GME_QUOTITE';
if TypeDEB = 'I' then LeNomSQL := LeNomSql+', PY_CODEBANCAIRE';
NuLibre := 0;
if TypeDeb = 'E' then
   begin
   if MontantLibreE.Text <> '<<Aucun>>' then
         NuLibre := strtoint (Copy (MontantLibreE.Text, length (MontantLibreE.Text), 1));
   end else begin
   if MontantLibreI.Text <> '<<Aucun>>' then
      NuLibre := strtoint (Copy (MontantLibreI.Text, length (MontantLibreI.Text), 1));
   end;
if NuLibre <> 0 then
   begin
   LeNomSQL := LeNomSQL+', GA_VALLIBRE'+IntToStr(NuLibre);
   end;
LeNomSQL := LeNomSql+' FROM LIGNENOMEN LEFT JOIN ARTICLE ON GA_ARTICLE=GLN_ARTICLE';
LeNomSQL := LeNomSQL+' LEFT JOIN MEA ON GME_QUALIFMESURE="POI" AND GME_MESURE=GA_QUALIFPOIDS';
If TypeDEB = 'I' then
   LeNomSQL := LeNomSQL+' LEFT JOIN PAYS ON PY_PAYS=GA_PAYSORIGINE';
LeNomSQL := LeNomSQL+' WHERE '+WhereNom;
TOBNomenc := TOB.Create('',Nil,-1);
QQ := OpenSql(LeNomSQL,true);
if not QQ.Eof then TOBNomenc.LoadDetailDB('','','',QQ,false)
else TOBNomenc := Nil;
Ferme(QQ);
end;

//===============================================================
// éclatement d'une nomenclature jusqu'au premier NC8 ou article
//===============================================================
Procedure TFAssistDEBC.Eclater(var TobEclat : TOB; TOBNomenc : TOB; NumLigne : integer; Niveau : Integer; Compose : string; Ordre : Integer; Qte : double);
var TOBN : TOB;
    TOBM : TOB;
    Niv : integer;
    Qten : double;
    TOBF : TOB;
    Poids : double;
    Pvht : double;
BEGIN
If TOBNomenc = Nil then exit;
Niv := Niveau + 1;
TOBN := TOBNomenc.FindFirst(['GLN_NUMLIGNE','GLN_NIVEAU','GLN_COMPOSE','GLN_ORDRECOMPO'],
     [Numligne,Niv,Compose,Ordre],false);
While TOBN <> Nil do
      begin
      Qten := TOBN.GetValue('GLN_QTE');
      Qten := Qten * Qte;
      if (TOBN.GetValue('GA_TYPEARTICLE') <> 'NOM') OR (TOBN.GetValue('GA_CODEDOUANIER') <> '') then
         begin
         if TOBEclat = Nil then TOBEclat := TOB.Create('',Nil,-1);
         TOBN.PutValue('GLN_QTE',Qten);
         Poids := TOBN.GetValue('GA_POIDSDOUA');
         Poids := Poids * Qten;
         TOBN.PutValue('GA_POIDSDOUA',Poids);
         Pvht := TOBN.GetValue('GA_PVHT');
         Pvht := Pvht * Qten;
         TOBN.PutValue('GA_PVHT',Pvht);
         TOBN.ChangeParent(TOBEclat,-1);
         end else begin
         TOBM := Nil;
         Eclater(TOBM,TOBNomenc,NumLigne,TOBN.GetValue('GLN_NIVEAU'),TOBN.GetValue('GLN_ARTICLE'),
             TOBN.GetValue('GLN_NUMORDRE'),Qten);
         if TOBM = Nil then
            begin
            TOBN.PutValue('GLN_QTE',Qten);
            Poids := TOBN.GetValue('GA_POIDSDOUA');
            Poids := Poids * Qten;
            TOBN.PutValue('GA_POIDSDOUA',Poids);
            Pvht := TOBN.GetValue('GA_PVHT');
            Pvht := Pvht * Qten;
            TOBN.PutValue('GA_PVHT',Pvht);
            TOBN.ChangeParent(TOBEclat,-1);
            end else begin
            While TOBM.Detail.count <> 0 do
                begin
                TOBF := TOBM.Detail[0];
                TOBF.ChangeParent(TOBEclat,-1);
                end;
            TOBM.free;
            TOBN.Free;
            end;
         end;
      TOBN := TOBNomenc.FindFirst(['GLN_NUMLIGNE','GLN_NIVEAU','GLN_COMPOSE','GLN_ORDRECOMPO'],
           [Numligne,Niv,Compose,Ordre],false);
      end;
end;
//======================================================
// récupération de l'éclatement dans la tob des lignes
//======================================================
Procedure TFAssistDEBC.TobEclatToTobLigne(TypeDEB : string; var TobPiece : TOB; TobLigne : TOB; TOBCompos : TOB; var NbLig : Integer);
var Prorata : double;
    Diviseur : double;
    Part : double;
    ind : integer;
    TOBC : TOB;
    TOBL : TOB;
    NumLibre : Integer;
BEGIN
Diviseur := TOBCompos.Somme('GA_PVHT',['GLN_NUMLIGNE'],[TobLigne.GetValue('GL_NUMLIGNE')],False, False);
for ind := 0 to TOBCompos.Detail.count - 1 do
    begin
    TOBC := TOBCompos.detail[ind];
    If ((TOBC.GetValue('GA_TYPEARTICLE') = 'MAR') or (TOBC.GetValue('GA_TYPEARTICLE') = 'NOM')) and (TOBC.GetValue('GA_PVHT') <> 0)  then
       begin
       TOBL := TOB.Create('',TobPiece,-1);
       TOBL.Dupliquer(TobLigne,false,false);
       TOBL.PutValue('GL_NATUREPIECEG',TOBC.GetValue('GLN_NATUREPIECEG'));
       TOBL.PutValue('GL_SOUCHE',TOBC.GetValue('GLN_SOUCHE'));
       TOBL.PutValue('GL_NUMERO',TOBC.GetValue('GLN_NUMERO'));
       TOBL.PutValue('GL_INDICEG',TOBC.GetValue('GLN_INDICEG'));
       TOBL.PutValue('GL_CODEDOUANE',TOBC.GetValue('GA_CODEDOUANIER'));
       TOBL.PutValue('GL_QTEFACT',TOBC.GetValue('GLN_QTE'));
       TOBL.PutValue('GL_NUMLIGNE',TOBC.GetValue('GLN_NUMLIGNE'));
       TOBL.PutValue('GL_ARTICLE',TOBC.GetValue('GLN_ARTICLE'));
       TOBL.PutValue('GL_NUMORDRE',TOBC.GetValue('GLN_NUMORDRE'));
       TOBL.PutValue('GL_REFARTTIERS','');
       TOBL.PutValue('GL_TOTALPOIDSDOUA',TOBC.GetValue('GA_POIDSDOUA'));
       TOBL.PutValue('GL_CODEARTICLE',TOBC.GetValue('GLN_CODEARTICLE'));
       TOBL.PutValue('GL_TYPEARTICLE',TOBC.GetValue('GA_TYPEARTICLE'));
       TOBL.PutValue('GL_REFARTBARRE',TOBC.GetValue('GA_REFARTBARRE'));
       TOBL.PutValue('GDE_CODEPOSTAL',TobLigne.GetValue('GDE_CODEPOSTAL'));
       TOBL.PutValue('GME_QUOTITE',TOBC.GetValue('GME_QUOTITE'));
       If TypeDEB = 'I' then
          TOBL.PutValue('PY_CODEBANCAIRE',TOBC.GetValue('PY_CODEBANCAIRE'));
       NumLibre := 0;
       if TypeDeb = 'E' then
       begin
       if MontantLibreE.Text <> '<<Aucun>>' then
         NumLibre := strtoint (Copy (MontantLibreE.Text, length (MontantLibreE.Text), 1));
       end else begin
       if MontantLibreI.Text <> '<<Aucun>>' then
         NumLibre := strtoint (Copy (MontantLibreI.Text, length (MontantLibreI.Text), 1));
         end;
       if NumLibre <> 0 then
          begin
          TOBL.PutValue('GA_VALLIBRE'+IntToStr(NumLibre),TOBC.GetValue('GA_VALLIBRE'+IntToStr(NumLibre)));
          end;
       Prorata := TobLigne.GetValue('GL_TOTALHT');
       Part := TOBC.GetValue('GA_PVHT');
       Prorata := (Prorata * Part) / Diviseur;
       TOBL.PutValue('GL_TOTALHT',Prorata);
       Prorata := TobLigne.GetValue('GL_TOTALHTDEV');
       Prorata := (Prorata * Part) / Diviseur;
       TOBL.PutValue('GL_TOTALHTDEV',Prorata);
       NbLig := NbLig+1;
       end;
    end;
end;
//================================================
// sélection des natures de pièce à traiter
// procédure à modifier plus tard en fonction de nouveaux
// éventuels paramètres dans PARPIECE
//================================================
Procedure TFAssistDEBC.WhereParPiece(TypeDEB : string; var TOBNat : TOB; var WhereNat : string);
var ind : Integer;
    TOBDP : TOB;
    TOBDN : TOB;
    Regime : string;
    TransacA : string;
    TransacB : string;
    Concerne : boolean;
    possk : integer;
    QteMoins : string;
    prem : boolean;
BEGIN
prem := true;
for ind := 0 to VH_GC.TOBParPiece.Detail.count - 1 do
    begin
    TOBDP := VH_GC.TOBParPiece.Detail[ind];
    Concerne := false;
    if TypeDEB = 'I' then
       begin
       if TOBDP.GetValue('GPP_NATUREPIECEG') = 'FF' then
          begin
          Concerne := true;
          Regime := '11';
          TransacA := '1';
          TransacB := '1';
          end;
       end;
    If TypeDEB = 'E' then
       begin
       if TOBDP.GetValue('GPP_NATUREPIECEG') = 'FAC' then
          begin
          Concerne := true;
          Regime := '21';
          TransacA := '1';
          TransacB := '1';
          end;
       If TOBDP.GetValue('GPP_ESTAVOIR') = 'X' then
          begin
          QteMoins := TOBDP.GetValue('GPP_QTEMOINS');
          possk := Pos('PHY',QteMoins);
          if possk <> 0 then
             begin
             Concerne := true;
             Regime := '21';
             TransacA := '2';
             TransacB := '1';
             end;
          end;
       end;
    if concerne then
       begin
       if TOBNat = Nil then TOBNat := TOB.Create('',Nil,-1);
       TOBDN := TOB.Create('',TOBNat,-1);
       TOBDN.AddChampSup('GPP_NATUREPIECEG',True);
       TOBDN.AddChampSup('Regime',True);
       TOBDN.AddChampSup('TransacA',True);
       TOBDN.AddChampSup('TransacB',True);
       TOBDN.PutValue('GPP_NATUREPIECEG',TOBDP.GetValue('GPP_NATUREPIECEG'));
       TOBDN.PutValue('Regime',Regime);
       TOBDN.PutValue('TransacA',TransacA);
       TOBDN.PutValue('TransacB',TransacB);
       if prem then WhereNat := '('
       else WhereNat := WhereNat + ' OR ';
       prem := false;
       WhereNat := WhereNat + 'GP_NATUREPIECEG="'+TOBDP.GetValue('GPP_NATUREPIECEG')+'"';
       end;
    end;
if prem = false then WhereNat := WhereNat+')';
end;

//====================================================================
// Ecriture du journal des évènements
//====================================================================
Procedure TFAssistDEBC.EnregistreEvenement;
var NumJnl : integer;
    TOBJnl : TOB;
    QQ : TQuery;
begin
Numjnl:=0 ;
TOBJnl:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnl.PutValue('GEV_TYPEEVENT', 'DEB');
TOBJnl.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnl.PutValue('GEV_DATEEVENT', Date);
TOBJnl.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
if Not QQ.EOF then Numjnl:=QQ.Fields[0].AsInteger ;
Ferme(QQ);
Inc(Numjnl) ;
TOBJnl.PutValue('GEV_NUMEVENT', Numjnl);
TOBJnl.PutValue('GEV_ETATEVENT', 'OK');
TOBJnl.PutValue('GEV_BLOCNOTE',ListRecap.Items.Text );
TOBJnl.InsertDB(Nil) ;
TOBJnl.Free;
end;

//====================================================================
// fonction  de formatage
//====================================================================
//-------------------------- ajouter des blancs jusqu'à la longueur désirée
Procedure TFAssistDEBC.FormatChaine(var Zone : string; lg : integer);
var Long : integer;
begin
long := Length(Zone);
setLength(Zone,lg);
While Long < lg do
      begin
      Long := Long+1;
      Zone[Long] := ' ';
      end;
end;

procedure TFAssistDEBC.BfermerClick(Sender: TObject);
begin
  inherited;
Close ;
end;

end.
