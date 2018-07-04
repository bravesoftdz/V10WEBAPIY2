unit Expecc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,{ Filtre SG6 15/11/04 FQ 14826,}
  StdCtrls, Hcompte, HRegCpte, Hctrls, ComCtrls, hmsgbox, Menus, HSysMenu,
  HTB97, Ent1, HEnt1, Mask,HPanel, UiUtil, dbTables, ExtCtrls,UObjFiltres,
  IniFiles, Utiltrans;

Procedure ExportEnCours(StArg : string='') ;

type
  TFExpECC = class(TForm)
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Msg: THMsgBox;
    Sauve: TSaveDialog;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    BFiltre: TToolbarButton97;
    PLibres: TTabSheet;
    TabSheet2: TTabSheet;
    HLabel2: THLabel;
    ST413EAR: THLabel;
    ST413EEP: THLabel;
    STCPTSOLDE: THLabel;
    FMethodeECTNE: TRadioGroup;
    FMethodeECTENR: TRadioGroup;
    OKTL: TCheckBox;
    OkAutreMP: TCheckBox;
    TT_TABLE0: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE4: THCpteEdit;
    T_TABLE9: THCpteEdit;
    TCPTEDEBUT: THLabel;
    TCPTEFIN: THLabel;
    LFile: THLabel;
    LFDateButoir: THLabel;
    Label1: TLabel;
    HLabel3: THLabel;
    HDateNeg1: THLabel;
    HDateNeg2: THLabel;
    CPTEDEBUT: THCpteEdit;
    CPTEFIN: THCpteEdit;
    FDateButoir: THCritMaskEdit;
    Format: THValComboBox;
    Masque: TEdit;
    DateNeg1: THCritMaskEdit;
    DateNeg2: THCritMaskEdit;
    CBExportDetail: TCheckBox;
    Bevel1: TBevel;
    Bevel3: TBevel;
    Bevel2: TBevel;
    SFDateEcr1: THLabel;
    SFDateModif1: THLabel;
    FDateEcr1: THCritMaskEdit;
    FDateModif1: THCritMaskEdit;
    SFDateModif2: THLabel;
    SFDateEcr2: THLabel;
    FDateEcr2: THCritMaskEdit;
    FDateModif2: THCritMaskEdit;
    FileName: THCritMaskEdit;
    OKPL: TCheckBox;
    FTimer: TTimer;
    BIni: TToolbarButton97;
    TCPTSOLDE: THCritMaskEdit;
    T411: THCritMaskEdit;
    T413EAR: THCritMaskEdit;
    T413EEP: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormatChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBExportDetailClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FTimerTimer(Sender: TObject);
    procedure BIniClick(Sender: TObject);
  private
    { Déclarations privées }
    NomFiltre, Commande : String ;
    ObjFiltre : TObjFiltre; //Gestion des Filtres SG6 15/11/04  FQ 14826
    function RenseigneArg(Commande : String) : Boolean;  // fiche 10561
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  TImpFic, ImpFicU1 ;

Procedure ExportEnCours (StArg : string='') ;
var FExpEcc:TFExpEcc ;
    PP : THPanel ;
    Composants : TControlFiltre; //SG6   Gestion des Filtes 15/11/04   FQ 14826
BEGIN
FExpEcc:=TFExpEcc.Create(Application) ;

//SG6 10/11/04 Gestion des Filtres FQ 14826
Composants.PopupF   := FExpEcc.POPF;
Composants.Filtres  := FExpEcc.FFILTRES;
Composants.Filtre   := FExpEcc.BFILTRE;
Composants.PageCtrl := FExpEcc.Pages;
FExpEcc.ObjFiltre   := TObjFiltre.Create(Composants,'');
FExpEcc.Commande    := stArg;

PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     if StArg = '' then
          FExpEcc.ShowModal
     else   // fiche 10561
     begin
                     FExpEcc.RenseigneArg (StArg);
                     FExpEcc.FTimerTimer(Application);
     end;
     finally
     FExpEcc.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FExpEcc,PP) ;
   FExpEcc.Show ;
   END ;
END ;

procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
var j,i : integer ;
BEGIN
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then BEGIN j:=i ; Break ; END ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
END ;

procedure InitTableLibre ( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : TLabel ;
    CC      : THCpteEdit ;
BEGIN
if ((EstSerie(S3)) or (EstSerie(S5))) then BEGIN TT.TabVisible:=False ; Exit ; END ;
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ; GetLibelleTableLibre('T',LesLib) ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
    LL:=TLabel(TForm(TT.Owner).FindComponent('TT_TABLE'+IntToStr(i))) ;
    if LL<>Nil then
       BEGIN
       LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
       CC:=THCpteEdit(LL.FocusControl) ;
       if CC<>Nil then BEGIN CC.Enabled:=(St='X') ; if CC.Enabled then Trouver:=True ; END ;
       END ;
    END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;


procedure TFExpECC.FormShow(Sender: TObject);
begin
Pages.ActivePage:=Pages.Pages[0] ;
NomFiltre:='TEST' ;
{$IFDEF SPEC302}
InitTableLibre(PLibres) ;
{$ENDIF}
FDateButoir.Text:=DateToStr(V_PGI.DateEntree) ;
DateNeg1.Text:=DateToStr(VH^.Entree.Deb) ; DateNeg2.Text:=DateToStr(VH^.Entree.Fin) ;
Format.ItemIndex:=0 ; FormatChange(Nil) ;

//SG6 15/11/04 Gestion des Filtres  FQ 14826
ObjFiltre.FFI_TABLE:=NomFiltre;
ObjFiltre.Charger;
FTimer.Enabled := FALSE;
  if Commande <> '' then  // par ligne de commande fiche 10561
  begin
         if not RenseigneArg(Commande) then exit;
  end;
end;

Procedure RAZFec(Var FEC : tFicEnCours) ;
BEGIN
VideStringList(FEC.LR) ; FEC.LR:=Nil ;
END ;

procedure TFExpECC.BValiderClick(Sender: TObject);
Var FEC : tFicEnCours ;
    Crit : tCritExpECC ;
    i : Integer ;
    LL : THCpteEdit ;
begin
if Commande = '' then
begin
  if Msg.Execute(0,Caption,'')<>mrYes then Exit ;
end;
If FileName.Text='' Then BEGIN Msg.Execute(2,Caption,'') ; Pages.ActivePage:=Pages.Pages[0] ; FileName.SetFocus ; Exit ; END ;
If Not CBExportDetail.Checked Then
  BEGIN
  If T411.Text='' Then BEGIN Msg.Execute(3,Caption,'') ; Pages.ActivePage:=Pages.Pages[1] ; T411.SetFocus ; Exit ; END ;
  If T413EEP.Text='' Then BEGIN Msg.Execute(4,Caption,'') ; Pages.ActivePage:=Pages.Pages[1] ; T413EEP.SetFocus ; Exit ; END ;
  END ;
EnableControls(Self,False) ;
Fillchar(FEC,SizeOf(FEC),#0) ; FillChar(Crit,SizeOf(Crit),#0) ;
Crit.Aux1:=CpteDebut.text ; Crit.Aux2:=CpteFin.text ;
Crit.St413EAR:=T413EAR.Text ; Crit.St413EEP:=T413EEP.Text ;
Crit.St411:=T411.Text ; Crit.Format:=Format.Value ;
Crit.DateButoir:=StrToDate(FDateButoir.Text) ;
Crit.DateNeg1:=StrToDate(DateNeg1.Text) ; Crit.DateNeg2:=StrToDate(DateNeg2.Text) ;
If FMethodeECTNE.ItemIndex=1 Then Crit.MethodeECTNE:=TraiteEtFacture ;
If FMethodeECTNE.ItemIndex=2 Then Crit.MethodeECTNE:=Global ;
If FMethodeECTENR.ItemIndex=1 Then Crit.MethodeECTENR:=TraiteEtFacture ;
Crit.St411SOLDE:=TCPTSOLDE.Text ;
Crit.OkTousMP:=OkAutreMP.Checked ; 
If CBExportDetail.Checked Then
  BEGIN
  Crit.DateCpta1:=StrToDate(FDateEcr1.Text) ; Crit.DateCpta2:=StrToDate(FDateEcr2.Text) ;
  Crit.DateModif1:=StrToDate(FDateModif1.Text) ; Crit.DateModif2:=StrToDate(FDateModif2.Text) ;
  Crit.OkPL:=OkPL.Checked ; Crit.OkTL:=OkTL.Checked ;
  END ;
Crit.Masque:=trim(Masque.Text) ;
For i:= 0 To 9 Do
  BEGIN
  LL:=THCpteEdit(Self.FindComponent('T_TABLE'+IntToStr(i))) ;
  if (LL<>Nil) And (LL.Text<>'') then Crit.TL[i]:=LL.Text ;
  END ;
initFEC(FEC) ; FEC.NomFic:=FileName.Text ; FEC.Format:=Crit.Format ;
If CBExportDetail.Checked Then
  BEGIN
  FaitFichierEnCoursDetail(FEC,Crit) ;
  END Else
  BEGIN
  FaitEnCoursComptable(FEC,Crit) ;
  FaitEnCoursTraiteNonEchu(FEC,Crit) ;
  If Crit.Format<>'NEG' Then FaitEnCoursTraiteEchuNonRegle(FEC,Crit) ;
  FaitFichierEnCours(FEC,Crit) ;
  END ;
RazFEC(FEc) ;
if Commande = '' then
    Msg.Execute(1,Caption,'') ;
EnableControls(Self,True) ;
end;


procedure TFExpECC.FormatChange(Sender: TObject);
begin
DateNeg1.Enabled:=Format.Value='NEG' ; DateNeg2.Enabled:=Format.Value='NEG' ;
HDateNeg1.Enabled:=Format.Value='NEG' ; HDateNeg2.Enabled:=Format.Value='NEG' ;
FMethodeECTENR.Enabled:=Format.Value<>'NEG' ;
CBExportDetail.enabled:=Format.Value='PRO' ;
OkAutreMP.Enabled:=Format.Value='ORL' ;
If Not CBExportDetail.enabled Then CBExportDetail.Checked:=FALSE ;
If Not OkAutreMP.enabled Then OkAutreMP.Checked:=FALSE ;
end;

procedure TFExpECC.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre);  //SG6 15/11/04 Gestion des Filtres FQ 14826
if Parent is THPanel then
  BEGIN
  Action:=caFree ;
  END ;
if Commande <> '' then
begin
    FTimer.free;
    ModalResult := 1;
    if WindowState  = wsMinimized then
       FiniProgressbar;
end;

end;

procedure TFExpECC.CBExportDetailClick(Sender: TObject);
begin
FMethodeECTNE.Visible:=Not CBExportDetail.Checked ; FMethodeECTENR.Visible:=Not CBExportDetail.Checked ;
FDateButoir.Enabled:=Not CBExportDetail.Checked ; LFDateButoir.Enabled:=Not CBExportDetail.Checked ;
T413EAR.Enabled:=Not CBExportDetail.Checked ; ST413EAR.Enabled:=Not CBExportDetail.Checked ;
T413EEP.Enabled:=Not CBExportDetail.Checked ; ST413EEP.Enabled:=Not CBExportDetail.Checked ;
TCPTSOLDE.Enabled:=Not CBExportDetail.Checked ; STCPTSOLDE.Enabled:=Not CBExportDetail.Checked ;
OkPL.Visible:=CBExportDetail.Checked ; OkTL.Visible:=CBExportDetail.Checked ;
FDateEcr1.Visible:=CBExportDetail.Checked ; SFDateEcr1.Visible:=CBExportDetail.Checked ;
FDateEcr2.Visible:=CBExportDetail.Checked ; SFDateEcr2.Visible:=CBExportDetail.Checked ;
FDateModif1.Visible:=CBExportDetail.Checked ; SFDateModif1.Visible:=CBExportDetail.Checked ;
FDateModif2.Visible:=CBExportDetail.Checked ; SFDateModif2.Visible:=CBExportDetail.Checked ;
end;

procedure TFExpECC.HMTradBeforeTraduc(Sender: TObject);
begin
{$IFNDEF SPEC302}
LibellesTableLibre(PLibres,'TT_TABLE','T_TABLE','T') ;
{$ENDIF}
end;

//Sg6 13/01/05 FQ 15242
procedure TFExpECC.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFExpECC.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //SG6 13/01/05 FQ 15242
  if key = VK_ESCAPE then BFermeClick(nil)
end;

// fiche 10561

function TFExpECC.RenseigneArg(Commande : String) : Boolean;
var
tmp, St            : string;
FicIni             : TIniFile;
stArgsav           : string;
               Procedure IniEnabled (Champ : string; CTR : TWinControl);
               begin
                  St := FicIni.ReadString ('COMMANDE', Champ, '');
                  TCheckBox(CTR).Checked := (St = 'TRUE') or (St = 'X');
               end;
begin
            Result := TRUE; stArgsav := Commande;
            if pos('Minimized', Commande) <> 0 then
                 WindowState  := wsMinimized;
            FTimer.Enabled := TRUE;
            tmp            := Commande;
            if (pos('.INI', Uppercase(tmp)) <> 0) then
            begin
                      tmp    := ReadTokenPipe(Commande, ';');
                      if not FileExists(tmp) then
                      begin
                          PGIInfo ('Le fichier '+ tmp+ ' n''existe pas','');
                          Result := FALSE;
                          exit;
                      end;
                      FicIni                   := TIniFile.Create(tmp);
                      tmp                      := FicIni.ReadString ('COMMANDE', 'FORMAT', '');
                      Format.Itemindex         := Format.Values.IndexOf(tmp);
                      Format.value             := tmp;
                      IniEnabled ('EXPORTDETAIL', CBExportDetail);

                      CPTEDEBUT.text           := FicIni.ReadString ('COMMANDE', 'AUXDEBUT', '');
                      CPTEFIN.text             := FicIni.ReadString ('COMMANDE', 'AUXFIN', '');
                      Masque.text              := FicIni.ReadString ('COMMANDE', 'MASQUE', '');
                      FDateButoir.Text         := FicIni.ReadString ('COMMANDE', 'DATEBUTOIR', '');
                      FileName.text            := FicIni.ReadString ('COMMANDE', 'NOMFICHIER', '');
                      DateNeg1.text            := FicIni.ReadString ('COMMANDE', 'DATENEG1', '');
                      DateNeg2.text            := FicIni.ReadString ('COMMANDE', 'DATENEG2', '');
                      TCPTSOLDE.text            := FicIni.ReadString ('COMMANDE', 'CPTESOLDECLI', '');
                      T411.Text                := FicIni.ReadString ('COMMANDE', 'CPTEFACTCLI', '');
                      T413EAR.text             := FicIni.ReadString ('COMMANDE', 'EFFETARECEVOIR', '');
                      T413EEP.text             := FicIni.ReadString ('COMMANDE', 'EFFETPORTEFEUILLE', '');
                      IniEnabled ('CALCULMDP', OkAutreMP);
                      St := FicIni.ReadString ('COMMANDE', 'TRAITENONECHU', '');
                      if (St = '1') or (St = '2') or (St = '3') then FMethodeECTNE.ItemIndex := StrToInt(St);

                      St := FicIni.ReadString ('COMMANDE', 'TRAITENONREGLE', '');
                      if (St = '1') or (St = '2')  then FMethodeECTENR.ItemIndex := StrToInt(St);

                      IniEnabled ('OKPL', OKPL);
                      IniEnabled ('OKTL', OKTL);
                      FDateEcr1.text            := FicIni.ReadString ('COMMANDE', 'DATE1', '');
                      FDateEcr2.text            := FicIni.ReadString ('COMMANDE', 'DATE2', '');
                      FDateModif1.text          := FicIni.ReadString ('COMMANDE', 'DATEMODIF1', '');
                      FDateModif2.text          := FicIni.ReadString ('COMMANDE', 'DATEMODIF2', '');

                      T_TABLE0.text             := FicIni.ReadString ('COMMANDE', 'TABLE1', '');
                      T_TABLE1.text             := FicIni.ReadString ('COMMANDE', 'TABLE2', '');
                      T_TABLE2.text             := FicIni.ReadString ('COMMANDE', 'TABLE3', '');
                      T_TABLE3.text             := FicIni.ReadString ('COMMANDE', 'TABLE4', '');
                      T_TABLE4.text             := FicIni.ReadString ('COMMANDE', 'TABLE5', '');
                      T_TABLE5.text             := FicIni.ReadString ('COMMANDE', 'TABLE6', '');
                      T_TABLE6.text             := FicIni.ReadString ('COMMANDE', 'TABLE7', '');
                      T_TABLE7.text             := FicIni.ReadString ('COMMANDE', 'TABLE8', '');
                      T_TABLE8.text             := FicIni.ReadString ('COMMANDE', 'TABLE9', '');
                      T_TABLE9.text             := FicIni.ReadString ('COMMANDE', 'TABLE10', '');
                      FicIni.free;
            end;
            Commande := stArgsav;
             if WindowState  = wsMinimized then
             begin
               St := 'Execution Export Fichier : ' + FileName.text;
               InitProgressbar(St);
             end;

end;

procedure TFExpECC.FTimerTimer(Sender: TObject);
begin
  FTimer.Enabled := FALSE;
  BValiderClick(Sender);
  Close ;
end;

procedure TFExpECC.BIniClick(Sender: TObject);
var
FicIni                   : TIniFile;
Rep                      : string;
Fbat                     : TextFile;
Commande,CurrentFile     : string;
Currentbat,FF            : string;
Dossier, St              : string;
ii                       : integer;
begin
  inherited;
      Rep := ExtractFileDir (FileName.Text);
      SetCurrentDirectory(PChar(Rep));
       with Topendialog.create(Self) do
       begin
         FileName := 'COMENCOURS.INI';
         Filter := 'Fichiers texte (*.INI)|*.INI|Tous les fichiers (*.*)|*.*';
         FilterIndex := 1;
         if Execute then
           CurrentFile := FileName
         else
            CurrentFile := '';
         Free;
       end;
      if  CurrentFile = '' then exit
      else
      Rep := ExtractFileDir (CurrentFile);
      Currentbat := CurrentFile;
      FF := ReadTokenPipe (Currentbat, '.');;
      Currentbat := FF +'.bat';


      if FileExists (CurrentFile) then DeleteFile( PChar(CurrentFile) );
      if FileExists (Currentbat) then DeleteFile( PChar(Currentbat) );

      FicIni        := TIniFile.Create(CurrentFile);
      FicIni.WriteString ('COMMANDE', 'FORMAT', Format.value);
      if CBExportDetail.Checked then
        FicIni.WriteString ('COMMANDE', 'EXPORTDETAIL',  'TRUE' )
      else
        FicIni.WriteString ('COMMANDE', 'EXPORTDETAIL',  'FALSE' );

        FicIni.WriteString ('COMMANDE', 'AUXDEBUT', CPTEDEBUT.text);
        FicIni.WriteString ('COMMANDE', 'AUXFIN',  CPTEFIN.text);
        FicIni.WriteString ('COMMANDE', 'MASQUE', Masque.text);
        FicIni.WriteString ('COMMANDE', 'DATEBUTOIR', FDateButoir.Text);
        FicIni.WriteString ('COMMANDE', 'NOMFICHIER', FileName.text);
        FicIni.WriteString ('COMMANDE', 'DATENEG1', DateNeg1.text);
        FicIni.WriteString ('COMMANDE', 'DATENEG2', DateNeg2.text);
        FicIni.WriteString ('COMMANDE', 'CPTESOLDECLI', TCPTSOLDE.text);
        FicIni.WriteString ('COMMANDE', 'CPTEFACTCLI', T411.Text);
        FicIni.WriteString ('COMMANDE', 'EFFETARECEVOIR', T413EAR.text);
        FicIni.WriteString ('COMMANDE', 'EFFETPORTEFEUILLE', T413EEP.text);
        if OkAutreMP.Checked then
        FicIni.WriteString ('COMMANDE', 'CALCULMDP',  'TRUE' )
      else
        FicIni.WriteString ('COMMANDE', 'CALCULMDP',  'FALSE' );
      if not CBExportDetail.Checked then
      begin
            Case (FMethodeECTNE.ItemIndex) of
            1 : FicIni.WriteString ('COMMANDE', 'TRAITENONECHU', '1' );
            2 : FicIni.WriteString ('COMMANDE', 'TRAITENONECHU', '2' );
            3 : FicIni.WriteString ('COMMANDE', 'TRAITENONECHU', '3' );
            end;
            Case FMethodeECTNE.ItemIndex of
            1 : FicIni.WriteString ('COMMANDE', 'TRAITENONREGLE', '1' );
            2 : FicIni.WriteString ('COMMANDE', 'TRAITENONREGLE', '2' );
            end;
      end
      else
      begin
            if OKPL.Checked then  FicIni.WriteString ('COMMANDE', 'OKPL', 'TRUE' );
            if OKTL.Checked then  FicIni.WriteString ('COMMANDE', 'OKPL', 'OKTL' );
            FicIni.WriteString ('COMMANDE', 'DATE1', FDateEcr1.text);
            FicIni.WriteString ('COMMANDE', 'DATE2', FDateEcr2.text);
            FicIni.WriteString ('COMMANDE', 'DATEMODIF1', FDateModif1.text);
            FicIni.WriteString ('COMMANDE', 'DATEMODIF2', FDateModif2.text);
      end;
      FicIni.WriteString ('COMMANDE', 'TABLE1', T_TABLE0.text);
      FicIni.WriteString ('COMMANDE', 'TABLE2', T_TABLE1.text);
      FicIni.WriteString ('COMMANDE', 'TABLE4', T_TABLE3.text);
      FicIni.WriteString ('COMMANDE', 'TABLE5', T_TABLE4.text);
      FicIni.WriteString ('COMMANDE', 'TABLE6', T_TABLE5.text);
      FicIni.WriteString ('COMMANDE', 'TABLE7', T_TABLE6.text);
      FicIni.WriteString ('COMMANDE', 'TABLE8', T_TABLE7.text);
      FicIni.WriteString ('COMMANDE', 'TABLE9', T_TABLE8.text);
      FicIni.WriteString ('COMMANDE', 'TABLE10', T_TABLE9.text);

      FicIni.free;

      // fichier .bat
    AssignFile(Fbat, Currentbat);
    Rewrite(Fbat) ;
    for ii :=1 to ParamCount do
    begin
        St:=ParamStr(ii) ;
        Dossier :=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
        if Dossier='/DOSSIER'  then begin Dossier := St; break; end;
    end;

    if ((V_PGI.ModePCL='1')) then
      Commande := Application.ExeName +' /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+Dossier+ ' /INI='+CurrentFile+';ENCOURS;Minimized'
    else
        Commande := Application.ExeName +' /USER='+V_PGI.UserLogin+ ' /PASSWORD='+V_PGI.PassWord+' /DOSSIER='+V_PGI.Currentalias+ ' /INI='+CurrentFile+';ENCOURS;Minimized';
    Writeln(Fbat, Commande) ;
    CloseFile(Fbat);
    PGIInfo ('La génération du fichier de commande est terminée');

end;

end.
