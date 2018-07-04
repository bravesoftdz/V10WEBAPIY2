{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/08/2004
Modifié le ... : 24/04/2007
Description .. : FQ 13023 - CA - 11/08/2004 : choix de la balance de
Suite ........ : situation indépendante de l'exercice déjà sélectionné
Suite ........ : FQ 13021 - CA - 11/08/2004 : ne pas autoriser le choix
Suite ........ : d'une situation si on choisit de travailler sur une formule
Suite ........ : FQ 12859 - CA - 11/08/2004 : on stoppe l'édition si on
Suite ........ : annule le choix des options d'impression (cas d'un état par
Suite ........ : section)
Suite ........ : CA - 24/04/2007 : On supprime le choix d'affichage des
Suite ........ : devises car non satisfaisant en édition : sera remplacé plus
Suite ........ : tard par une devise de reporting. En attendant, on peut
Suite ........ : réactiver la zone avec le code spécif 51214
Mots clefs ... : 
*****************************************************************}
unit CRITSYNTH;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Mask,
    HTB97,
    StdCtrls,
    ExtCtrls,
    Hctrls,
    ComCtrls,
    CritEdt,
    Menus,
    HSysMenu,
    hmsgbox,
    HPanel,
    UiUtil,
    Filtre,
    Ent1,
    HEnt1,
    ParamDat,
    Spin,
    Buttons,
    utob,
    HZoomSp,
{$IFDEF EAGLCLIENT}
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    BalSit,
    UListByUser,
    UTXML,
    Math,
    TabLiRub,
    ParamSoc,
    ZCumul,
    uYFILESTD, TntButtons, TntStdCtrls, TntExtCtrls;

procedure LanceEtatSynthese(TE:TTypeEtatSynthese;vCritEdtChaine:TCritEdtChaine);
procedure EtatPCL(TE:TTypeEtatSynthese);
procedure EdtSynthCRChaine(vCritEdtChaine:TCritEdtChaine);
procedure EdtSynthSIGChaine(vCritEdtChaine:TCritEdtChaine);
procedure EdtSynthBILChaine(vCritEdtChaine:TCritEdtChaine);

type
    TEtatPCL = class(TForm)
        Pages:TPageControl;
        Standards:TTabSheet;
        Dock971:TDock97;
        PanelFiltre:TToolWindow97;
        FFiltres:THValComboBox;
        HPB:TToolWindow97;
        BParamListe:TToolbarButton97;
        BValider:TToolbarButton97;
        BFerme:TToolbarButton97;
        BAide:TToolbarButton97;
        GB1:TGroupBox;
        Exo1:THValComboBox;
        FD11:TMaskEdit;
        FD12:TMaskEdit;
        OkCol1:TCheckBox;
        Complements:TTabSheet;
        HLabel7:THLabel;
        HLabel8:THLabel;
        FDevises:THValComboBox;
        FEtab:THValComboBox;
        FAffiche:TRadioGroup;
        GB2:TGroupBox;
        HLabel2:THLabel;
        Label1:TLabel;
        Exo2:THValComboBox;
        FD21:TMaskEdit;
        FD22:TMaskEdit;
        OkCol2:TCheckBox;
        GB3:TGroupBox;
        HLabel9:THLabel;
        Label2:TLabel;
        Exo3:THValComboBox;
        FD31:TMaskEdit;
        FD32:TMaskEdit;
        OkCol3:TCheckBox;
        GB4:TGroupBox;
        HLabel11:THLabel;
        Label3:TLabel;
        Exo4:THValComboBox;
        FD41:TMaskEdit;
        FD42:TMaskEdit;
        OkCol4:TCheckBox;
        FAvec:TGroupBox;
        FSimu:TCheckBox;
        FSitu:TCheckBox;
        FRevi:TCheckBox;
        MsgBox:THMsgBox;
        HMTrad:THSystemMenu;
        POPF:TPopupMenu;
        BCreerFiltre:TMenuItem;
        BSaveFiltre:TMenuItem;
        BDelFiltre:TMenuItem;
        BRenFiltre:TMenuItem;
        BNouvRech:TMenuItem;
        BFiltre:TToolbarButton97;
        Sauve:TSaveDialog;
        FVoirNom:TCheckBox;
        Type1:THValComboBox;
        Type2:THValComboBox;
        Type3:THValComboBox;
        Type4:THValComboBox;
        Formule1:TEdit;
        Formule2:TEdit;
        Formule3:TEdit;
        Formule4:TEdit;
        TFRESOL:THLabel;
        FRESOL:THValComboBox;
        HLabel5:THLabel;
        FFormat:TEdit;
        FVariation:TRadioGroup;
        FValVar1:TSpinEdit;
        FValVar2:TSpinEdit;
        FValVar3:TSpinEdit;
        FValVar4:TSpinEdit;
        FMontant0:TCheckBox;
        PFormat:TPanel;
        H_TitreGuide:TLabel;
        PFenGuide:TPanel;
    BCValide: THBitBtn;
    BCAbandon: THBitBtn;
        CBMP:TCheckBox;
        FMP:TEdit;
        EMP:TLabel;
        CBMN:TCheckBox;
        FMN:TEdit;
        EMN:TLabel;
        CBMU:TCheckBox;
        FMU:TEdit;
        EMU:TLabel;
        tFInfoLibre:THLabel;
        FInfoLibre:TEdit;
        HLabel13:THLabel;
        Histobal1:THCritMaskEdit;
        Histobal2:THCritMaskEdit;
        Histobal3:THCritMaskEdit;
        Histobal4:THCritMaskEdit;
        AvecSitu1:TCheckBox;
        AvecSitu2:TCheckBox;
        AvecSitu3:TCheckBox;
        AvecSitu4:TCheckBox;
        HLabel1:THLabel;
        GroupBox1:TGroupBox;
        HLabel12:THLabel;
        HLabel14:THLabel;
        HLabel15:THLabel;
        HLabel16:THLabel;
        HLabel17:THLabel;
        HLabel18:THLabel;
        HLabel4:THLabel;
        FPrevi:TCheckBox;
        Tab_Analytique:TTabSheet;
        rdg_Analytique:THRadioGroup;
        HLabel3:THLabel;
        FNatureCpt:THValComboBox;
        pnl_Comptes:TPanel;
        HLabel6:THLabel;
        HLabel10:THLabel;
        edt_Section:TEdit;
        edt_SectionEx:TEdit;
        pnl_TLibres:TPanel;
        TFaG:TLabel;
        TFGen:THLabel;
        TFGenJoker:THLabel;        
        FSansANO:TCheckBox;
        HLabel19:THLabel;
        FRESOLDETAIL:THValComboBox;
        HLabel20:THLabel;
        FFormatDetail:TEdit;
        PBilan:TTabSheet;
        GBBILAN:TGroupBox;
        BILEXO:THValComboBox;
        BILDATE:TMaskEdit;
        BILDATE_:TMaskEdit;
        BILBALSIT:THCritMaskEdit;
        BILAVECSIT:TCheckBox;
        HLabel21:THLabel;
        HLabel22:THLabel;
        HLabel23:THLabel;
        GBBILANCOMP:TGroupBox;
        BILEXOCOMP:THValComboBox;
        BILDATECOMP:TMaskEdit;
        BILDATECOMP_:TMaskEdit;
        BILBALSITCOMP:THCritMaskEdit;
        BILAVECSITCOMP:TCheckBox;
        HLabel25:THLabel;
        HLabel26:THLabel;
        HLabel27:THLabel;
        BILCOMP:TCheckBox;
        FSansCompte:TCheckBox;
        Panel1:TPanel;
        LFile:THLabel;
        ZMAQUETTENAME:THLabel;
        ZMAQUETTE:THCritMaskEdit;
        Panel2:TPanel;
        HLabel24:THLabel;
        ZMAQUETTENAMEBIL:THLabel;
        ZMAQUETTEBIL:THCritMaskEdit;
        FAvecPourcent:TCheckBox;
        Timer1:TTimer;
        HLabel28:THLabel;
        FMentionFiligrane:TEdit;
        FUnEtatParSection:TCheckBox;
        FLibelleDansLesTitres:TCheckBox;
        Dupliquerlefiltre1:TMenuItem;
        FIFRS: TCheckBox;
        FLibelleSection: TCheckBox;
        FImpressionDate: TCheckBox;
        FAvecDetail: THValComboBox;
        FAvecDetailBil: THValComboBox;
        BExport: TToolbarButton97;
        FSECTION: THCritMaskEdit;
        FSECTION_: THCritMaskEdit;
    lCORRESPONDANCE: THLabel;
        FCORRESPONDANCE: THValComboBox;
        procedure BCreerFiltreClick(Sender:TObject);
        procedure BSaveFiltreClick(Sender:TObject);
        procedure BDelFiltreClick(Sender:TObject);
        procedure BRenFiltreClick(Sender:TObject);
        procedure BNouvRechClick(Sender:TObject);
        procedure Dupliquerlefiltre1Click(Sender:TObject);
        procedure FormClose(Sender:TObject;var Action:TCloseAction);
        procedure FormCreate(Sender:TObject);
        procedure FormShow(Sender:TObject);
        procedure OkCol1Click(Sender:TObject);
        procedure Exo1Change(Sender:TObject);
        procedure BValiderClick(Sender:TObject);
        procedure BParamListeClick(Sender:TObject);
        procedure FD11KeyPress(Sender:TObject;var Key:Char);
        procedure Type1Change(Sender:TObject);
        procedure Formule1Enter(Sender:TObject);
        procedure Formule1Exit(Sender:TObject);
        procedure FVariationClick(Sender:TObject);
        procedure FD11Exit(Sender:TObject);
        procedure FRESOLChange(Sender:TObject);
        procedure FRESOLDetailChange(Sender:TObject);
        procedure CBMPClick(Sender:TObject);
        procedure FMPChange(Sender:TObject);
        procedure BCValideClick(Sender:TObject);
        procedure BCAbandonClick(Sender:TObject);
        procedure FFormatKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
        procedure FFormatDetailKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
        procedure OnClickAvecSituation(Sender:TObject);
        procedure OnBalSitClick(Sender:TObject);
        procedure ZMAQUETTEChange(Sender:TObject);
        procedure FAfficheClick(Sender:TObject);
        procedure BAideClick(Sender:TObject);
        procedure FNatureCptChange(Sender:TObject);
        procedure edt_SectionDblClick(Sender:TObject);
        procedure edt_SectionKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
        procedure rdg_AnalytiqueClick(Sender:TObject);
        procedure BILEXOChange(Sender:TObject);
        procedure BILAVECSITClick(Sender:TObject);
        procedure BILCOMPClick(Sender:TObject);
        procedure Timer1Timer(Sender:TObject);
        procedure FFiltresChange(Sender:TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word;
          Shift: TShiftState);
        procedure BFermeClick(Sender: TObject);
        procedure BExportClick(Sender: TObject);
    procedure FSECTIONChange(Sender: TObject);
    procedure FSECTIONKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FSECTIONKeyPress(Sender: TObject; var Key: Char);
    private
    { Déclarations privées }
        Crit:TCritEdtPCL;
        TE:TTypeEtatSynthese;
        FNomFiltre:string;
        IndF:Integer;
        fFichierMaquette:string;
        gbChargementFiltre:boolean;
        fbFocusFormatDetail:boolean;
        FCritEdtChaine:TCritEdtChaine;
        FListeByUser:TListByUser;
        procedure ActiveCol(GB:tGroupBox;OkOk:Boolean);
        function RecupCrit:Boolean;
        procedure SwapCol(GB:tGroupBox;OnRub:Boolean);
        procedure AjoutePourcentCrit;
        procedure RecalcFVariation(indice:Integer;Valeur:shortint);
        procedure ChangeFormat(var St:string;Avec:Boolean);
        procedure SwapEcranFormat(OkOk:Boolean);
        procedure InitEcranFormat(bDetail:boolean);
        function GetFichierMaquette:string;
        procedure EnableCritCol(Col:integer);
        function RecupCritBilan:Boolean;
//        procedure InitChoixMonnaie;
        procedure InitEtatChaine;
        // Code du MUL pour la nouvelle Gestion des FILTRES
        procedure InitAddFiltre(T:TOB);
        procedure InitGetFiltre(T:TOB);
        procedure InitSelectFiltre(T:TOB);
        procedure ParseParamsFiltre(Params: HTStrings);
        procedure UpgradeFiltre(T:TOB);
        procedure LanceEdition(pbExport: boolean);
        function  ControleCritere ( var vStMsg : string ) : boolean;
        procedure InitComboPlanCorrespondance;
//        procedure InitRadioGroupDevise;
    public
    { Déclarations publiques }
    end;

implementation

uses
  {$IFDEF MODENT1}
    CPTypeCons,
    CPProcMetier,
    CPProcGen,
    CPVersion,
  {$ENDIF MODENT1}
    GridSYNTH,
{$IFDEF EAGLCLIENT}
{$ELSE}
    UtilEdt,
{$ENDIF}
    uLibExercice
    ;

{$R *.DFM}

////////////////////////// Lancement états de synthèse                    /////////////////////////////

procedure EtatPCL(TE:TTypeEtatSynthese);
var
    CritEdtChaine:TCritEdtChaine;
begin
    Fillchar(CritEdtChaine, SizeOf(CritEdtChaine), #0);
    CritEdtChaine.Utiliser := False;
    LanceEtatSynthese(TE, CritEdtChaine);
end;

////////////////////////// Lancement états de synthèse chainés            /////////////////////////////

procedure EdtSynthCRChaine(vCritEdtChaine:TCritEdtChaine);
begin
    LanceEtatSynthese(esCR, vCritEdtChaine);
end;

procedure EdtSynthSIGChaine(vCritEdtChaine:TCritEdtChaine);
begin
    LanceEtatSynthese(esSIG, vCritEdtChaine);
end;

procedure EdtSynthBILChaine(vCritEdtChaine:TCritEdtChaine);
begin
    LanceEtatSynthese(esBIL, vCritEdtChaine);
end;

procedure ImporteMaquetteEtatSynth;
var
  lQ : TQuery;
  lStFichierMaquette, lStCheminMaquette : string;
  lRet : integer;
  stFileExt : string;
begin
  { Si il existe déjà des états de synthèse personnalisés en base on ne fait pas l'import }
  if ExisteSQL ('SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND YFS_CRIT1="ETATSYNTH" AND YFS_PREDEFINI="STD"') then
  begin
    SetParamSoc('SO_CPETATSYNTHESEENBASE',True);
    exit;
  end;
  try
    lQ := OpenSQL ('SELECT * FROM STDMAQ WHERE STM_NUMPLAN>20', True);
    try
      while not lQ.Eof do
      begin
        lStFichierMaquette := lQ.FindField('STM_TYPEMAQ').AsString + Format('%.03d', [lQ.FindField('STM_NUMPLAN').AsInteger]) + '.TXT';
        lStCheminMaquette := ChangeStdDatPath('$DAT\'+lStFichierMaquette,True);
        stFileExt := ExtractFileExt (lStCHeminMaquette);
        if (Pos('.',stFileExt)=1) then stFileExt := Copy (stFileExt,2, Length(stFileExt)-1);
        lRet := AGL_YFILESTD_IMPORT(lStCheminMaquette,'COMPTA',ExtractFileName(lStCheminMaquette), stFileExt,
                'ETATSYNTH',
                lQ.FindField('STM_TYPEMAQ').AsString,
                lQ.FindField('STM_NUMPLAN').AsString,
                '','','-','-','-','-','-',
                V_PGI.LanguePrinc,
                'STD',
                lQ.FindField('STM_LIBELLE').AsString);
        if lRet <> -1 then PGIInfo (AGL_YFILESTD_GET_ERR(lRet) + #13#10 + lStCheminMaquette);
        lQ.Next;
      end;
    finally
      Ferme (lQ);
    end;
    SetParamSoc('SO_CPETATSYNTHESEENBASE',True);
  except
  end;
end;

////////////////////////// Lancement général des états de synthèse        /////////////////////////////

procedure LanceEtatSynthese(TE:TTypeEtatSynthese;vCritEdtChaine:TCritEdtChaine);
var
    XX:TEtatPCL;
    PP:THPanel;
    i:Integer;
(*    T : TOB;
    bEntete : boolean;
    stEncoding : string;*)
begin
      // A DETRUIRE
//    TEST_GETCUMUL;
    // TEST XML
//    T :=TOB.Create ('',nil,-1);
//    T.LoadFromXMLFile ('c:\test.xml',bEntete,stEncoding);
//  T.Free;
    { Importation en base des états de synthèse paramétrés par l'utilisateur }
    if ctxPCL in V_PGI.PGIContexte then
      if not GetParamSocSecur('SO_CPETATSYNTHESEENBASE',True) then ImporteMaquetteEtatSynth;

    XX := TEtatPCL.Create(Application);
    XX.FCritEdtChaine := vCritEdtChaine;
    i := -1;
    XX.ZMAQUETTENAME.Caption := '';
    XX.ZMAQUETTENAMEBIL.Caption := '';
    if not (ctxPCL in V_PGI.PGIContexte) then
    begin
        XX.ZMAQUETTENAME.Visible := False;
        XX.ZMAQUETTE.Width := 150;
//        XX.ZMAQUETTE.DataType := 'OPENFILE(*.txt)';
        XX.ZMAQUETTE.DataType := 'CPETATSYNTH';
        case TE of
          esSIG, esSIGA : XX.ZMAQUETTE.Plus := ' AND YFS_CRIT2="SIG"';
          esCR, esCRA : XX.ZMAQUETTE.Plus := ' AND YFS_CRIT2="CR"';
        end;
        XX.ZMAQUETTENAMEBIL.Visible := False;
        XX.ZMAQUETTEBIL.Width := 150;
//        XX.ZMAQUETTEBIL.DataType := 'OPENFILE(*.txt)';
        XX.ZMAQUETTEBIL.DataType := 'CPETATSYNTH';
        XX.ZMAQUETTEBIL.Plus := ' AND YFS_CRIT2="BIL"'
    end;
    XX.PBilan.TabVisible := (TE = esBil) or (TE = esBILA);
    XX.Standards.TabVisible := not ((TE = esBil) or (TE = esBILA));
    XX.HelpContext := 999999802;
    case TE of
        esSIG:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTE.Plus := 'SIG';
                XX.FNomFiltre := 'ETATPCLSIG';
                i := 6;
            end;
        esBIL:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTEBIL.Plus := 'BIL';
                XX.FNomFiltre := 'ETATPCLBIL';
                i := 4;
            end;
        esCR:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTE.Plus := 'CR';
                XX.FNomFiltre := 'ETATPCLCR';
                i := 5;
            end;
        esCRA:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTE.Plus := 'CR';
                XX.FNomFiltre := 'ETATCRANA';
                i := 7;
                VH^.bAnalytique := True;
            end;
        esSIGA:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTE.Plus := 'SIG';
                xx.FNomFiltre := 'ETATSIGA';
                i := 8;
                VH^.bAnalytique := True;
            end;
        esBILA:
            begin
                if ctxPCL in V_PGI.PGIContexte then XX.ZMAQUETTEBIL.Plus := 'BIL';
                XX.FNomFiltre := 'ETATBILA';
                i := 9;
                VH^.bAnalytique := True;
            end;
    end;
    if i > 0 then XX.Caption := XX.MsgBox.Mess[0] + ' - ' + XX.MsgBox.Mess[i];
    XX.TE := TE;
    XX.Crit.TE := TE;
  //XX.HelpContext:=7802000 ;
    PP := FindInsidePanel;
    if (PP = nil) or (XX.FCritEdtChaine.Utiliser) then
    begin
        try
            if (XX.FCritEdtChaine.Utiliser) then XX.Timer1.Enabled := True;
            XX.ShowModal;
        finally
            XX.Free;
        end;
    end
    else
    begin
        if ctxPCL in V_PGI.PGIContexte then XX.AutoSize := False;
        InitInside(XX, PP);
        XX.Show;
    end;
end;

procedure TEtatPCL.ActiveCol(GB:tGroupBox;OkOk:Boolean);
var
    I:Integer;
    ChildControl:TControl;
begin
    if GB = nil then Exit;
    for I := 0 to GB.ControlCount - 1 do
    begin
        ChildControl := GB.Controls[I];
        if ChildControl <> nil then
        begin
            if ChildControl.Tag = 0 then ChildControl.Enabled := OkOk;
        end;
    end;
end;

procedure TEtatPCL.SwapCol(GB:tGroupBox;OnRub:Boolean);
var
    I:Integer;
    ChildControl:TControl;
begin
    if GB = nil then Exit;
    for I := 0 to GB.ControlCount - 1 do
    begin
        ChildControl := GB.Controls[I];
        if ChildControl <> nil then
        begin
            if (Pos('Type', ChildControl.Name) = 0) and (ChildControl.Tag = 0) then
            begin
        //      If Pos('Formule',ChildControl.Name)>0 Then ChildControl.Visible:=Not onRub Else ChildControl.Visible:=OnRub ;
                if Pos('Formule', ChildControl.Name) > 0 then
                    ChildControl.Enabled := not onRub
                else ChildControl.Enabled := OnRub;
            end;
        end;
    end;
end;

procedure TEtatPCL.FFiltresChange(Sender:TObject);
var
    i:Integer;
    HIV:tHvalComboBox;
    lExoDate:TExoDate;
    FD1:TMaskEdit;
    FD2:TMaskEdit;
begin
    gbChargementFiltre := True;
    FListeByUser.Select(FFiltres.Value);
    if not PBilan.TabVisible then
    begin
        for i := 1 to 4 do
        begin
            HIV := THvalComboBox(FindComponent('Exo' + IntToStr(i)));
            FD1 := tMaskEdit(FindComponent('FD' + IntToStr(i) + '1'));
            if FD1 = nil then
                Exit;
            FD2 := tMaskEdit(FindComponent('FD' + IntToStr(i) + '2'));
            if FD2 = nil then
                Exit;
            if HIV <> nil then
            begin
                lExoDate.Code := CRelatifVersExercice(HIV.Value);
                RempliExoDate(lExoDate);
                if (IsValidDate(FD1.Text) and IsValidDate(FD2.Text)) then
                begin
                    if ((StrToDate(FD1.Text) < lExoDate.Deb) or
                        (StrToDate(FD1.Text) > lExoDate.Fin)) or
                        ((StrToDate(FD2.Text) < lExoDate.Deb) or
                        (StrToDate(FD2.Text) > lExoDate.Fin)) then
                        Exo1Change(HIV);
                end;
            end;
        end;
    end
    else
    begin
        lExoDate.Code := CRelatifVersExercice(BILEXO.Value);
        RempliExoDate(lExoDate);
        if (IsValidDate(BILDATE.Text) and IsValidDate(BILDATE_.Text)) then
        begin
            if ((StrToDate(BILDATE.Text) < lExoDate.Deb) or
                (StrToDate(BILDATE.Text) > lExoDate.Fin)) or
                ((StrToDate(BILDATE_.Text) < lExoDate.Deb) or
                (StrToDate(BILDATE_.Text) > lExoDate.Fin)) then
                BILEXOChange(BILEXO);
        end;
        if BILCOMP.Checked then
        begin
            lExoDate.Code := CRelatifVersExercice(BILEXOCOMP.Value);
            RempliExoDate(lExoDate);
            if (IsValidDate(BILDATECOMP.Text) and IsValidDate(BILDATECOMP_.Text)) then
            begin
                if ((StrToDate(BILDATECOMP.Text) < lExoDate.Deb) or
                    (StrToDate(BILDATECOMP.Text) > lExoDate.Fin)) or
                    ((StrToDate(BILDATECOMP_.Text) < lExoDate.Deb) or
                    (StrToDate(BILDATECOMP_.Text) > lExoDate.Fin)) then
                BILEXOChange(BILEXOCOMP);
            end;
        end;
    end;
    gbChargementFiltre := False;
    if fAvecDetail.Value = '' then FAvecDetail.Value := 'NON';
    if fAvecDetailBil.Value = '' then FAvecDetailBil.Value := 'NON';
end;

procedure TEtatPCL.BCreerFiltreClick(Sender:TObject);
begin
    FListeByUser.Creer;
    FListeByUser.Save;
end;

procedure TEtatPCL.BSaveFiltreClick(Sender:TObject);
begin
    FListeByUser.Save;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TEtatPCL.Dupliquerlefiltre1Click(Sender:TObject);
begin
    FListeByUser.Duplicate;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TEtatPCL.BDelFiltreClick(Sender:TObject);
begin
    VideFiltre(FFiltres, Pages);
    FListeByUser.Delete;
    FListeByUser.Save;
end;

procedure TEtatPCL.BRenFiltreClick(Sender:TObject);
begin
    FListeByUser.Rename;
end;

procedure TEtatPCL.BNouvRechClick(Sender:TObject);
begin
    VideFiltre(FFiltres, Pages);
    FListeByUser.New;
end;

procedure TEtatPCL.FormClose(Sender:TObject;var Action:TCloseAction);
var
  bCanClose : Boolean;
begin
    if FListeByUser <> nil then
    begin
        if FListeByUser.IsModified then
            FListeByUser.CloseQuery(bCanClose);
        FreeAndNil(FListeByUser); // Libération du TListByUser
    end;
    GridPCL := False;
    VH^.bAnalytique := False; // Pour le compte de résultat et SIG analytique
    if Parent is THPanel then Action := caFree;
end;

procedure TEtatPCL.FormCreate(Sender:TObject);
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
end;

procedure TEtatPCL.FormShow(Sender:TObject);
var
    i:Integer;
    HIV:tHvalComboBox;
    SP:TSpinEdit;
    CB:TCheckBox;
    lExoDate:TExoDate;
    FD1:TMaskEdit;
    FD2:TMaskEdit;
begin
    GridPCL := TRUE;
    FAffiche.Visible := EstSpecif ('51214');

    { Initialisation des combos exercices }
    CInitComboExercice(Exo1);
    CInitComboExercice(Exo2);
    CInitComboExercice(Exo3);
    CInitComboExercice(Exo4);
    CInitComboExercice(BILEXO);
    CInitComboExercice(BILEXOCOMP);
    { Initialisation des exercices }
    BILEXO.Value := CExerciceVersRelatif (GetCPExoRef.Code);
    BILEXOCOMP.Value := CExerciceVersRelatif (GetCPExoRef.Code);    

    FAvecDetail.Value := 'NON';
    FAvecDetailBil.Value := 'NON';

    {$IFDEF CCSTD }
    bExport.Visible := False;
    {$ENDIF}

    { Activation onglet analytique }
    if VH^.bAnalytique then Tab_Analytique.TabVisible := True
    else Tab_Analytique.TabVisible := False;
    FLibelleSection.Visible := VH^.bAnalytique;

  // Activation/désactivation onglet spécifique bilan
    if PBilan.TabVisible = True then Pages.ActivePage := PBilan
    else Pages.ActivePage := Standards;
  // Initialisation Axe n°1
    if VH^.bAnalytique then
    begin
      FNatureCpt.ItemIndex := 0;
      FSECTION.DataType := 'TZSECTION';
      FSECTION_.DataType := 'TZSECTION';
    end;
  // Initialisation de la résolution
    FRESOL.Value := 'C';
    FResolChange(nil);
    FRESOLDETAIL.Value := 'C';
    FResolDetailChange(nil);
  // Initialisation des colonnes
    OkCol1.Checked := TRUE;
    OkCol2.Checked := FALSE;
    OkCol3.Checked := FALSE;
    OkCol4.Checked := FALSE;
    OkCol1Click(OkCol1);
    OkCol1Click(OkCol2);
    OkCol1Click(OkCol3);
    OkCol1Click(OkCol4);
    FAffiche.ItemIndex := 0;
    FInfoLibre.Text := GetNomCabinet;

    for i := 1 to 4 do
    begin
        HIV := THvalComboBox(FindComponent('Type' + IntToStr(i)));
        if HIV <> nil then HIV.Value := 'RUB';
        SP := TSpinEdit(FindComponent('FValVar' + IntToStr(i)));
        if SP <> nil then SP.Value := 0;
        CB := TCheckBox(FindComponent('AvecSitu' + IntToStr(i)));
        if CB <> nil then
            CB.Enabled := not VH^.bAnalytique;
        HIV := THvalComboBox(FindComponent('Histobal' + IntToStr(i)));
        if HIV <> nil then
            HIV.Enabled := not VH^.bAnalytique;
    end;
    gbChargementFiltre := True;
  // GCO - 04/12/2003
  // Création du TlisteByUser pour la nouvelle gestion des filtres
    FListeByUser := TListByUser.Create(FFiltres, BFiltre, toFiltre, False);
    with FListeByUser do
    begin
        OnSelect := InitSelectFiltre;
        OnInitGet := InitGetFiltre;
        OnInitAdd := InitAddFiltre;
        OnUpgrade := UpgradeFiltre;
        OnParams := ParseParamsFiltre;
    end;
    if FListeByUser <> nil then
      FListeByUser.LoadDB(FNomFiltre);

    FFiltres.Enabled := not (FCritEdtChaine.Utiliser);

    if FCritEdtChaine.Utiliser then InitEtatchaine;
    gbChargementFiltre := False;

    { Initialisation des comptes de correspondance }
    InitComboPlanCorrespondance;

  // Fin de la nouvelle gestion des filtres
    if FFormat.Text = '' then FFormat.Text := '#,##'+FaitMasqueDecimales+';; ;';
    if FFormatDetail.Text = '' then FFormatDetail.Text := '#,##'+FaitMasqueDecimales+';; ;';
    FAvecPourcent.Visible := TE <> esBIL;
    if ctxPCL in V_PGI.PGIContexte then
    begin
        FInfoLibre.Visible := True;
        tFInfoLibre.Visible := True;
    end;

    FIfrs.Visible :=  VH^.OkModCPPackIFRS;
    if FIFRS.Visible then FAvec.Height := FAvec.Height+10;

    GBBILANCOMP.Enabled := BILCOMP.Checked;

    if not PBilan.TabVisible then
    begin
        for i := 1 to 4 do
        begin
            HIV := THvalComboBox(FindComponent('Exo' + IntToStr(i)));
            FD1 := tMaskEdit(FindComponent('FD' + IntToStr(i) + '1'));
            if FD1 = nil then
                Exit;
            FD2 := tMaskEdit(FindComponent('FD' + IntToStr(i) + '2'));
            if FD2 = nil then
                Exit;
            if HIV <> nil then
            begin
                lExoDate.Code := CRelatifVersExercice(HIV.Value);
                RempliExoDate(lExoDate);
                if (IsValidDate(FD1.Text) and IsValidDate(FD2.Text)) then
                begin
                    if ((StrToDate(FD1.Text) < lExoDate.Deb) or
                        (StrToDate(FD1.Text) > lExoDate.Fin)) or
                        ((StrToDate(FD2.Text) < lExoDate.Deb) or
                        (StrToDate(FD2.Text) > lExoDate.Fin)) then
                        Exo1Change(HIV);
                end;
            end;
        end;
    end
    else
    begin
        lExoDate.Code := CRelatifVersExercice(BILEXO.Value);
        RempliExoDate(lExoDate);
        if (IsValidDate(BILDATE.Text) and IsValidDate(BILDATE_.Text)) then
        begin
            if ((StrToDate(BILDATE.Text) < lExoDate.Deb) or
                (StrToDate(BILDATE.Text) > lExoDate.Fin)) and
                ((StrToDate(BILDATE_.Text) < lExoDate.Deb) or
                (StrToDate(BILDATE_.Text) > lExoDate.Fin)) then
                BILEXOChange(BILEXO);
        end;
        if BILCOMP.Checked then
        begin
            lExoDate.Code := CRelatifVersExercice(BILEXOCOMP.Value);
            RempliExoDate(lExoDate);
            if (IsValidDate(BILDATECOMP.Text) and IsValidDate(BILDATECOMP_.Text)) then
            begin
                if ((StrToDate(BILDATECOMP.Text) < lExoDate.Deb) or
                    (StrToDate(BILDATECOMP.Text) > lExoDate.Fin)) and
                    ((StrToDate(BILDATECOMP_.Text) < lExoDate.Deb) or
                    (StrToDate(BILDATECOMP_.Text) > lExoDate.Fin)) then
                    BILEXOChange(BILEXOCOMP);
            end;
        end;
    end;

{$IFDEF CCSTD}
  FListeByUser.ForceAccessibilite := FaPublic;
{$ENDIF}
  PositionneEtabUser(FEtab, False); // 15088
end;

procedure TEtatPCL.OkCol1Click(Sender:TObject);
var
    CB:tCheckBox;
    GB:TGroupBox;
    HIV:THValComboBox;
    SP : TSpinEdit;
begin
    CB := TCheckBox(Sender);
    if CB = nil then Exit;
    if CB.Tag > 0 then
    begin
        GB := TGroupBox(FindComponent('GB' + IntToStr(CB.Tag)));
        if GB <> nil then ActiveCol(GB, CB.Checked);
    end;
    EnableCritCol(CB.Tag);
    HIV := THvalComboBox(FindComponent('Type' + IntToStr(CB.Tag)));
    FVariation.Visible := (FVariation.Visible) and ((CB.Checked) and (HIV.Value = 'FOR'));
    CB := TCheckBox(FindComponent('AvecSitu' + IntToStr(CB.Tag)));
    if CB <> nil then
        CB.Enabled := not VH^.bAnalytique;
    HIV := THvalComboBox(FindComponent('Histobal' + IntToStr(CB.Tag)));
    if HIV <> nil then
        HIV.Enabled := not VH^.bAnalytique;
    SP := TSpinEdit(FindComponent('FValVar' + IntToStr(CB.Tag)));
    if SP <> nil then SP.Value := 0;
end;

procedure TEtatPCL.Exo1Change(Sender:TObject);
var
    HIV:THValComboBox;
    GB:tGroupBox;
    FD1, FD2:tMaskEdit;
    HB:THCritMaskEdit;
begin
    HIV := THValComboBox(Sender);
    if HIV = nil then Exit;
    GB := TGroupBox(HIV.Parent);
    if GB = nil then Exit;
    FD1 := tMaskEdit(FindComponent('FD' + IntToStr(GB.Tag) + '1'));
    if FD1 = nil then Exit;
    FD2 := tMaskEdit(FindComponent('FD' + IntToStr(GB.Tag) + '2'));
    if FD2 = nil then Exit;
    CExoRelatifToDates(HIV.Value, FD1, FD2);
  // Si changement d'exercice et balance de situation sélectionnée, on 'vide' la balance
    if not gbChargementFiltre then // CA - 29/06/2004 - FQ 13598 - On ne remet pas la balance à jour en cas de chargement d'un filtre.
    begin
      HB := THCritMaskEdit(FindComponent('Histobal' + IntToStr(GB.Tag)));
      if HB = nil then Exit;
      HB.Text := '';
      HB.Hint := '';
    end;
end;

function ConvStUneDate(d:tDateTime):string;
var
    YY, MM, DD:Word;
begin
    DecodeDate(d, YY, MM, DD);
    Result := '(' + FormatFloat('00', DD) + '-' + FormatFloat('00', MM) + '-' + FormatFloat('0000', YY) + ')';
end;

function QuelTitre(Date1, Date2:tDateTime):string;
var
    YY2, MM2, DD2, YY1, MM1, DD1:Word;
    Deb, Fin:Boolean;
    St:string;
    stDuree:string;
    PremMois, PremAnnee, NbMois:Word;
begin
    DecodeDate(Date2, YY2, MM2, DD2);
    DecodeDate(Date1, YY1, MM1, DD1);
    St := '';
    Deb := (DebutDeMois(Date1) = Date1);
    Fin := (FinDeMois(Date2) = Date2);
    if ctxPCL in V_PGI.PGIContexte then
    begin
        NOMBREMOIS(Date1, Date2, PremMois, PremAnnee, NbMois);
        if NbMois > 0 then stDuree := IntToStr(NbMois) + TraduireMemoire(' mois')
        else stDuree := FormatFloat('#', Date2 - Date1) + TraduireMemoire(' jours'); // Inutile, jamais affiché.
        // Inutile, jamais affiché.
    end;
    if (MM1 = MM2) and (YY1 = YY2) then
    begin
        if Deb and Fin then
        begin
            St := FormatDateTime('MMMM YYYY', Date1);
        end
        else
        begin
            St := FormatDateTime('MMM YYYY', Date1) + #13 + #10 + '(' + FormatFloat('00', DD1) + '-' + FormatFloat('00', DD2) + ')';
        end;
    end
    else
    begin
        if Deb and Fin then
        begin
      {      if ctxPCL in V_PGI.PGIContexte then
              St:=FormatDateTime('MMM YYYY',Date1)+#13+#10+FormatDateTime('MMM YYYY',Date2)
            else } St := TraduireMemoire('du ') + FormatDateTime('dd/mm/yy',
                Date1) + #13 + #10 + TraduireMemoire('au ') +
                FormatDateTime('dd/mm/yy', Date2) + #13 + #10 + stDuree;
        end
        else
        begin
      {      if ctxPCL in V_PGI.PGIContexte then
              St:=FormatDateTime('DD-MMM-YY',Date1)+#13+#10+FormatDateTime('DD-MMM-YY',Date2)
            else } St := TraduireMemoire('du ') + FormatDateTime('dd/mm/yy',
                Date1) + #13 + #10 + TraduireMemoire('au ') +
                FormatDateTime('dd/mm/yy', Date2) + #13 + #10 + stDuree;
        end;
    end;
    Result := St;
end;

function ConvStDate(var C:tColPCL):string;
var
    i :Integer;
    st:string;
    iEncours : integer;
begin
  // St:='-' ;
    St := ''; // Correction CA le 15/05/2001 - Sinon décalage d'1 exercice
    if (C.Exo.Deb = C.Date1) and (C.Exo.Fin = C.Date2) then
    begin
        if C.Exo.Code = VH^.EnCours.Code then
        begin
            Result := 'N';
            C.StTitre := 'N';
        end
        else if C.Exo.Code = VH^.Precedent.Code then
        begin
            Result := 'N-';
            C.StTitre := 'N-1';
        end
        else if C.Exo.Code = VH^.Suivant.Code then
        begin
            Result := 'N+';
            C.StTitre := 'N+1';
        end
        else
        begin
          i :=1;
          while (VH^.Exercices[i].Code<>VH^.Encours.Code) do
            Inc (i) ;
          iEncours := i;
          while (VH^.Exercices[i].Code<>C.Exo.Code) do
          begin
            St := St + '-';
            Dec (i);
          end;
          if ((i>0) and (VH^.Exercices[i].Code=C.Exo.Code)) then
          begin
            Result := 'N' + St;
            C.StTitre := 'N-' + IntToStr(iEncours-i);
          end;
        end;
        St := ConvStUneDate(C.Date1);
        St := St + ConvStUneDate(C.Date2);
        C.StTitre := QuelTitre(C.Date1, C.Date2);
    end
    else
    begin
        St := ConvStUneDate(C.Date1);
        St := St + ConvStUneDate(C.Date2);
        Result := St;
        C.StTitre := QuelTitre(C.Date1, C.Date2);
    end;
end;

procedure DecaleIndice(var St:string);
var
    St1, St2:string;
    i:Integer;
    C:Char;
begin
    if St = '' then Exit;
    St1 := St;
    i := Pos('COL', ST1);
    while i > 0 do
    begin
        St1 := FindEtReplace(ST1, 'COL', 'WWW', FALSE);
        St2 := Copy(St1, i + 3, 1);
        C := St2[1];
        case C of
            '2':C := '3';
            '3':C := '5';
            '4':C := '7';
        end;
        St[i + 3] := C;
        i := Pos('COL', ST1);
    end;
end;

function PourCentSurFormule(St:string;l:Integer):string;
var
    St1, St2, St3:string;
    k:Integer;
begin
    k := Pos('-', St);
    if k <= 0 then
    begin
        Result := '=100*[COL' + IntToStr(l) + ']/[BASE' + IntToStr(l) + ']';
        Exit;
    end;
    St1 := Copy(St, k + 1, Length(St) - k);
    St2 := Copy(St, 2, Length(St) - 1);
    St3 := '=100*(' + St2 + ')/(' + St1 + ')';
    DecaleIndice(St3);
    Result := St3;
end;

procedure TEtatPCL.AjoutePourcentCrit;
var
    CritMirror:tCritEdtPCL;
    i:Integer;
begin
    CritMirror := Crit;
    for i := 1 to 4 do
        if CritMirror.Col[i - 1].Actif then
        begin
            Crit.Col[2 * i - 2] := CritMirror.Col[i - 1];
            Crit.Col[2 * i - 1] := CritMirror.Col[i - 1];
            with Crit.Col[2 * i - 1] do
            begin
                Actif := TRUE;
                Fillchar(Exo, SizeOf(Exo), #0);
                Date1 := 0;
                Date2 := 0;
                AvecHisto := FALSE;
                IdentHisto := '';
                if ctxPCL in V_PGI.PGIContexte then
                begin
                    StFormat := '#,##0.00;; ;';
                    StFormatDetail := '#,##0.00;; ;';
                end
                else
                begin
                    StFormat := Crit.Col[2 * i - 2].StFormat;
                    StFormatDetail := Crit.Col[2 * i - 2].StFormatDetail;
                end;
                if Crit.Col[2 * i - 2].IsFormule = FALSE then StFormule := '=100*[COL' + IntToStr(2 * i - 1) + ']/[BASE' + IntToStr(2 * i - 1) + ']'
                else StFormule := PourcentSurFormule(Crit.Col[2 * i - 2].StFormule, 2 * i - 1);
                IsFormule := TRUE;
                if ctxPCL in V_PGI.PGIContexte then StTitre := '%'
                else StTitre := 'Pourcentage';
            end;
            with Crit.Col[2 * i - 2] do
                if Crit.Col[2 * i - 2].Actif then
                begin
                    if IsFormule then DecaleIndice(StFormule);
                end;
            Inc(Crit.NbColActif);
        end;
end;

procedure DecomposeNomFic(NomFic:string;var LePath, LeNom, Lextension:string);
begin
    LePath := ExtractFilePath(NomFic);
    LExtension := ExtractFileExt(NomFic);
    LeNom := ExtractFileName(NomFic);
end;

function TEtatPCL.RecupCrit:Boolean;
var
    i, j:Integer;
    CodeExo:string;
    HIV:tHValComboBox;
    OnRub:Boolean;
    SP:TSpinEdit;
    stSQL : string;
    Q : TQuery;
    bRet : boolean;
    stMsg : string;
    OkCol : TCheckBox;
begin
    Result := FALSE;
    OnRub := False;
    Fillchar(Crit, SizeOf(Crit), #0);
    Crit.bUnEtatParSection := False;
    Crit.TE := TE;
    fFichierMaquette := GetFichierMaquette;
    if fFichierMaquette = '' then
    begin
        MsgBox.execute(1, Caption, '');
        Exit;
    end;
    bRet :=  ControleCritere (stMsg);
    if (not bRet) then exit;
    Result := TRUE;
    if (TE = esBIL) or (TE = esBILA) then RecupCritBilan
    else
    begin
        for i := 1 to 4 do
        begin
            HIV := THvalComboBox(FindComponent('Type' + IntToStr(i)));
            if HIV <> nil then OnRub := HIV.Value = 'RUB';
            OkCol := TCheckBox(FindComponent('OkCol' + IntToStr(i)));
            if OkCol = nil then continue;
            if not OnRub AND (OkCol.Checked) then
            begin
                SP := TSpinEdit(FindComponent('FValVar' + IntToStr(i)));
                if SP <> nil then
                begin
                  j := SP.Value;
                  RecalcFVariation(i, j);
                end;
            end;
        end;
        for i := 1 to 4 do
            with Crit do
            begin
                Col[i - 1].Actif := TCheckBox(FindComponent('OkCol' + IntToStr(i))).Checked;
                if Col[i - 1].Actif then
                begin
                    HIV := THvalComboBox(FindComponent('Type' + IntToStr(i)));
                    if HIV <> nil then OnRub := HIV.Value = 'RUB';
                    if OnRub then
                    begin
                        CodeExo := CRelatifVersExercice(THvalComboBox(FindComponent('Exo' +
                            IntToStr(i))).Value);
                        QuelDateDeExo(CodeExo, Col[i - 1].Exo);
                        Col[i - 1].Date1 := StrToDate(TMaskEdit(FindComponent('FD' + IntToStr(i) + '1')).Text);
                        Col[i - 1].Date2 := StrToDate(TMaskEdit(FindComponent('FD' + IntToStr(i) + '2')).Text);
            // CA - 12/10/2001
                        Col[i - 1].BalSit := THCritMaskEdit(FindComponent('Histobal' + IntToStr(i))).Text;
            // CA - 12/10/2001 }
                        Col[i - 1].IsFormule := FALSE;
                    end
                    else
                    begin
                        Col[i - 1].StFormule := TEdit(FindComponent('Formule' + IntToStr(i))).Text;
                        Col[i - 1].IsFormule := TRUE;
                        SP := TSpinEdit(FindComponent('FValVar' + IntToStr(i)));
                        if SP <> nil then
                            case SP.Value of
                                0:Col[i - 1].StTitre := 'Simple : ' + #13 + #10 + 'Variation en valeur';
                                1:Col[i - 1].StTitre := 'Sur 12 mois : ' + #13 + #10 + 'Variation en valeur annuelle';
                                2:Col[i - 1].StTitre := 'Variation relative ' + #13 + #10 + 'au plus petit';
                                3:Col[i - 1].StTitre := 'Variation relative ' + #13 + #10 + 'au plus grand';
                            end;
                    end;
                end;
            end;
    end;
    with Crit do
    begin
        AvecPourcent := FAvecPourcent.Checked;
        if (TE = esBIL) or (TE = esBILA) then
        begin
          AvecDetail := (FAvecDetailBil.Value = 'SEUL') or (FAvecDetailBil.Value = 'PLUS');
          bEtat := (FAvecDetailBil.Value <> 'SEUL');
        end else
        begin
          AvecDetail := (FAvecDetail.Value = 'SEUL') or (FAvecDetail.Value = 'PLUS');
          bEtat := (FAvecDetail.Value <> 'SEUL');
        end;
        Modele := fFichierMaquette;
        if FEtab.ItemIndex > 0 then ETab := FEtab.Value;
        if FDevises.ItemIndex > 0 then Devise := FDevises.Value;
        TypEcr[Reel] := TRUE;
        TypEcr[Simu] := FSimu.Checked;
        TypEcr[Situ] := FSitu.Checked;
        TypEcr[Revi] := FRevi.Checked;
        TypEcr[Previ] := FPrevi.Checked;
        TypEcr[Ifrs]  := FIFRS.Checked;      // Ajout IFRS 05/05/2004
        StTypEcr := '';
        if TypEcr[Reel] then StTypEcr := StTypEcr + 'N';
        if TypEcr[Simu] then StTypEcr := StTypEcr + 'S';
        if TypEcr[Situ] then StTypEcr := StTypEcr + 'U';
        if TypEcr[Revi] then StTypEcr := StTypEcr + 'R';
        if TypEcr[Previ] then StTypEcr := StTypEcr + 'P';
        if TypEcr[IFRS]  then StTypEcr := StTypEcr + 'I';            // Ajout IFRS 05/05/2004
        if (FSansANO.Checked) then StTypEcr := StTypEcr + '-';
        Resolution := FResol.Value;
        ResolutionDetail := FResolDetail.Value;
        AvecMontant0 := FMontant0.Checked;
        bLibelleCompte := FSansCompte.Checked;
        bLibelleSection := FLibelleSection.Checked;
        InfoLibre := FInfoLibre.Text;
        MentionFiligrane := FMentionFiligrane.Text;
        DeviseEnPivot := '-';
        EnMonnaieOpposee := '-';
        if FAffiche.ItemIndex = 1 then DeviseEnPivot := 'X';
        bImpressionDate := FImpressionDate.Checked;
        Corresp := FCORRESPONDANCE.Value;
    end;
    for i := 1 to 4 do
        with Crit do
            if Col[i - 1].Actif then
            begin
                Inc(NbColActif);
                if Col[i - 1].StFormule = '' then Col[i - 1].StFormule := ConvStDate(Col[i - 1]);
        //    if Col[i-1].BalSit <> '' then Col[i-1].StTitre := Col[i-1].StTitre+#13+#10+'- Situation -';
            { CA - 15/04/2003 - Suite demande JPA, on n'affiche pas la mention situation }
                if Col[i - 1].BalSit <> '' then
                    Col[i - 1].StTitre := Col[i - 1].StTitre;
                if FFormat.Text <> '' then
                    Col[i - 1].StFormat := FFormat.Text
                else
                    Col[i - 1].StFormat := '#,##'+FaitMasqueDecimales+'';
                if FFormatDetail.Text <> '' then
                    Col[i - 1].StFormatDetail := FFormatDetail.Text
                else
                    Col[i - 1].StFormatDetail := '#,##'+FaitMasqueDecimales+'';
            end;
    if Crit.AvecPourcent then AjoutePourcentCrit;
    if VH^.bAnalytique then
    begin
        if (FNatureCpt.Value = '') then
        begin
            MsgBox.execute(11, Caption, '');
            Exit;
        end;
        if (rdg_Analytique.ItemIndex = 0) then
        begin
            Crit.bJoker := EstJoker(FSECTION.Text);
            Crit.stAxe := FNatureCpt.Value;
            if (Crit.bJoker) then
            begin
                Crit.stSectionDe := FSECTION.Text;
                // Crit.stSectionDe := FindEtReplace(Crit.stSectionDe, '*', '%', True);
            end
            else
            begin
              if (Trim(FSECTION.Text) = '') and (Trim(FSECTION_.Text) = '') then
              begin
                StSql := 'SELECT MIN(S_SECTION) MINI, MAX(S_SECTION) MAXI FROM SECTION ' +
                  'WHERE S_AXE="'+Crit.stAxe+'"';
                Q := OpenSQL (stSQL, True);
                if not Q.Eof then
                begin
                  FSECTION.Text := Q.FindField('MINI').AsString;
                  FSECTION_.Text := Q.FindField('MAXI').AsString;
                end;
              end;
              if (FSECTION.Text = '') then FSECTION.Text := FSECTION_.Text;
              Crit.stSectionDe := FSECTION.Text;
              if FSECTION_.Text = '' then FSECTION_.Text := FSECTION.Text;
              Crit.stSectionA := FSECTION_.Text;
            end;
            Crit.bUnEtatParSection := FUnEtatParSection.Checked;
            Crit.bLibelleDansLesTitres := FLibelleDansLesTitres.Checked;
        end
        else
        begin
            Crit.bTLI := True;
            Crit.stAxe := FNatureCpt.Value;
            if edt_Section.text = '' then edt_Section.text := '***,***,***,***,***,***,***,***,***,****';
            Crit.stSectionDe := edt_Section.text;
            Crit.stSectionA := edt_SectionEx.text;
        end;
    end;
    Result := TRUE;
end;

procedure TEtatPCL.BValiderClick(Sender:TObject);
begin
  LanceEdition ( False );
end;

procedure TEtatPCL.BParamListeClick(Sender:TObject);
begin
    Crit.bExport := False;
    fFichierMaquette := GetFichierMaquette;
    if FileExists(fFichierMaquette) then
        LanceLiasse(fFichierMaquette, FALSE, FVoirNom.Checked, ['N'], ['#,##'+FaitMasqueDecimales+''], ['#,##'+FaitMasqueDecimales+''], Crit, TRUE, FCritEdtChaine)
    else PGIInfo('Maquette introuvable !', Caption);
end;

procedure TEtatPCL.FD11KeyPress(Sender:TObject;var Key:Char);
begin
    ParamDate(Self, Sender, Key);
end;

procedure TEtatPCL.Type1Change(Sender:TObject);
var
    HIV:THValComboBox;
    GB:TGroupBox;
    Indice:integer;
    OkCol:TCheckBox;
    AvecSitu:TCheckBox;
    HSIT:THCritMaskEdit;
    Exo : THValComboBox;
begin
    HIV := THValComboBox(Sender);
    if HIV = nil then Exit;
    GB := TGroupBox(HIV.Parent);
    if GB = nil then Exit;
  // CA - 29/12/2001 - SwapCol déclenché uniquement pour les colonnes activées
  // - posait problème lors du chargement de la fiche
    Indice := StrToInt(Copy(HIV.Name, Length(HIV.Name), 1));
    OkCol := TCheckBox(FindComponent('OkCol' + IntToStr(Indice)));
    if OkCol.Checked then SwapCol(GB, HIV.Value = 'RUB');
    // CA - 26/04/2006 - Si choix RUBRIQUE alors on initialise l'exercice avec l'exercice de référence
    if assigned (HIV) then
    begin
      if HIV.Value = 'RUB' then
      begin
        Exo := THValComboBox(FindComponent('Exo' + IntToStr(Indice)));
        if assigned (Exo) then
        begin
          if Exo.Value = '' then Exo.Value := CExerciceVersRelatif (GetCPExoRef.Code);
        end;
      end;
    end;

  // Lien AvecSitu/Balsit
    AvecSitu := TCheckBox(FindComponent('AvecSitu' + IntToStr(Indice)));
    if AvecSitu <> nil then
    begin
        AvecSitu.Enabled := ((not VH^.bAnalytique) and (HIV.Value<>'FOR'));
        HSIT := THCritMaskEdit(FindComponent('Histobal' + IntToStr(Indice)));
        HSIT.Enabled := AvecSitu.Checked;
        { Si type = Formule, on ne peut pas choisir une situation - FQ 13021 }
        if HIV.Value = 'FOR' then
        begin
          AvecSitu.Checked := False;
          HSIT.Text := '';
        end;
    end;
end;

procedure TEtatPCL.Formule1Enter(Sender:TObject);
var
    Ed:TEdit;
    GB:tGroupBox;
    LeLeft, LeTop:Integer;
    SP:TSpinEdit;
begin
    Ed := TEdit(Sender);
    if Ed = nil then Exit;
    IndF := StrToInt(Copy(Ed.Name, Length(Ed.Name), 1));
    GB := tGroupBox(ED.Parent);
    LeTop := GB.Top + Ed.Top + Ed.Height + Standards.Top + 2;
    LeLeft := Ed.Left + GB.Left + Standards.Left;
    if LeTop + FVariation.Height > Pages.Height then LeTop := GB.Top + Ed.Top + Standards.Top - FVariation.Height - 2;

    FVariation.Top := LeTop;
    FVariation.Left := LeLeft;
    FVariation.Visible := TRUE;
  //TEdit(FindComponent('Formule'+IntToStr(IndF))).SetFocus ;
      SP := TSpinEdit(FindComponent('FValVar' + IntToStr(IndF)));
    if SP <> nil then FVariation.ItemIndex := SP.Value;
  { Ajout CA le 30/05/2001 - pour pb sur Win 98 suite mode inside }
    if not FVariation.Visible then FVariation.Visible := True;
    FVariation.BringToFront;
  { Fin Ajout CA }
end;

procedure TEtatPCL.Formule1Exit(Sender:TObject);
var
    Ac:TWincontrol;
begin
    AC := Self.ActiveControl;
    if Ac = nil then Exit;
    if (AC.Parent <> nil) and (AC.Parent.Name = FVariation.Name) then else FVariation.Visible := FALSE;
end;

function NBP(St1, St2:string):Word;
var
    Ex:tExoDate;
    XX, YY, ZZ:Word;
    Date1, Date2:tDateTime;
begin
    Date1 := StrToDate(St1);
    Date2 := StrToDate(St2);
    FillChar(Ex, SizeOf(Ex), #0);
    EX.Deb := Date1;
    EX.Fin := Date2;
    NOMBREPEREXO(Ex, XX, YY, ZZ);
    Result := ZZ;
end;

procedure TEtatPCL.FSECTIONKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  St : string;
  fb : TFichierBase ;
begin
  St := THCritMaskEdit(Sender).Text;
  fb := AxeToFb(FNatureCpt.Value);
  if (Shift = []) and (Key = 187) then
  begin
    Key := 0;
    FSECTION_.Text := FSECTION.Text;
  end
  else
  if ((Shift=[ssCtrl]) And (Key=VK_F5)) then
  begin
    if (fb in [fbAxe1..fbAxe5]) and
       VH^.Cpta[fb].Structure and
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"') then
    begin
      if ChoisirSousPlan( fb, St , True,taModif) then
       THCritMaskEdit(Sender).Text := St;
    end
    else
      THCritMaskEdit(Sender).ElipsisClick (Sender);
  end
  else
    if (Key=VK_F5) then
      THCritMaskEdit(Sender).ElipsisClick (Sender);

inherited KeyDown(Key,Shift) ;

end;

procedure TEtatPCL.RecalcFVariation(indice:Integer;Valeur:shortint);
var
    Ed:TEdit;
    i1, i2:Integer;
    j:Integer;
    HIV:THValComboBox;
    OnRub:Boolean;
    NbP1, nbP2:Word;
    NbpMin, NbpMax:Word;
    St:string;
    SP:TSpinEdit;
begin
    Ed := TEdit(FindComponent('Formule' + IntToStr(Indice)));
    i1 := 0;
    i2 := 0;
    j := 1;
    repeat
        HIV := THvalComboBox(FindComponent('Type' + IntToStr(j)));
        if HIV <> nil then
        begin
            OnRub := HIV.Value = 'RUB';
            if OnRub then
            begin
                if i1 = 0 then i1 := j else i2 := j;
            end;
        end;
        Inc(J);
    until ((i1 > 0) and (i2 > 0)) or (j > 4);
    if (i1 = 0) or (i2 = 0) then Exit;
    nbP1 := NBP(TMaskEdit(FindComponent('FD' + IntToStr(i1) + '1')).Text, TMaskEdit(FindComponent('FD' + IntToStr(i1) + '2')).Text);
    nbP2 := NBP(TMaskEdit(FindComponent('FD' + IntToStr(i2) + '1')).Text, TMaskEdit(FindComponent('FD' + IntToStr(i2) + '2')).Text);
    NbPMin := Nbp1;
    NbPMax := Nbp2;
    if NbP1 > NbP2 then
    begin
        NbPMin := Nbp2;
        NbPMax := Nbp1;
    end;
    case Valeur of
        0:
            begin
                St := '=[COL' + IntToStr(i1) + ']-[COL' + IntToStr(i2) + ']';
            end;
        1:
            begin
                St := '=(12*[COL' + IntToStr(i1) + ']/' + IntToStr(Nbp1) + ')-(12*[COL' + IntToStr(i2) + ']/' + IntToStr(Nbp2) + ')';
            end;
        2:
            begin
                St := '=(' + IntToStr(NbpMin) + '*[COL' + IntToStr(i1) + ']/' +
                    IntToStr(Nbp1) + ')-(' + IntToStr(NbpMin) + '*[COL' + IntToStr(i2) + ']/'
                    + IntToStr(Nbp2) + ')';
            end;
        3:
            begin
                St := '=(' + IntToStr(NbpMax) + '*[COL' + IntToStr(i1) + ']/' +
                    IntToStr(Nbp1) + ')-(' + IntToStr(NbpMax) + '*[COL' + IntToStr(i2) + ']/'
                    + IntToStr(Nbp2) + ')';
            end;
    end;
    Ed.Text := St;
    SP := TSpinEdit(FindComponent('FValVar' + IntToStr(Indice)));
    if SP <> nil then SP.Value := Valeur;
end;

procedure TEtatPCL.FVariationClick(Sender:TObject);
var
    Ed:TEdit;
begin
    RecalcFVariation(IndF, FVariation.ItemIndex);
    Ed := TEdit(FindComponent('Formule' + IntToStr(IndF)));
    Ed.SetFocus;
    FVariation.Visible := False;
end;

procedure TEtatPCL.FD11Exit(Sender:TObject);
var
    i, j:Integer;
    HIV:tHValComboBox;
    OnRub:Boolean;
    SP:TSpinEdit;
begin
    for i := 1 to 4 do
    begin
        HIV := THvalComboBox(FindComponent('Type' + IntToStr(i)));
        if HIV <> nil then
        begin
            OnRub := HIV.Value = 'RUB';
            if not OnRub then
            begin
                SP := TSpinEdit(FindComponent('FValVar' + IntToStr(i)));
                if SP <> nil then
                begin
                    j := SP.Value;
                    RecalcFVariation(i, j);
                end;
            end;
        end;
    end;
end;

procedure TEtatPCL.ChangeFormat(var St:string;Avec:Boolean);
var
    i, j, k:Integer;
    St1, St2:string;
  // '#,##0.00'
begin
    if not Avec then
    begin
        i := Pos('.', St);
        while i > 0 do
        begin
            k := Length(St);
            for j := i to Length(St) do
                if St[j] = ';' then
                begin
                    k := j - 1;
                    Break;
                end;
            Delete(st, i, k - i + 1);
            i := Pos('.', St);
        end;
    end
    else
    begin
        if Trim(St) = '' then Exit;
        i := Pos(';', St);
        if i = 0 then
            St := st + '.00'
        else
        begin
            St1 := St;
            St := '';
            while St1 <> '' do
            begin
                St2 := ReadTokenSt(St1);
                if (Trim(St2) <> '') and (Pos('.', St2) = 0) then St2 := St2 + '.00';
                St := St + St2 + ';';
            end;
        end;
    end;
end;

procedure TEtatPCL.FRESOLChange(Sender:TObject);
var
    St:string;
begin
    St := FFormat.text;
    if St = '' then Exit;
    if FResol.Value = 'C' then ChangeFormat(St, TRUE) else ChangeFormat(St, FALSE);
    FFormat.text := St;
end;

procedure TEtatPCL.FRESOLDetailChange(Sender:TObject);
var
    St:string;
begin
    St := FFormatDetail.text;
    if St = '' then Exit;
    if FResolDetail.Value = 'C' then ChangeFormat(St, TRUE) else ChangeFormat(St, FALSE);
    FFormatDetail.text := St;
end;

procedure TEtatPCL.CBMPClick(Sender:TObject);
var
    CB:TCheckBox;
    T, i:Integer;
    ChildControl:TControl;
begin
    CB := TCheckBox(Sender);
    if CB = nil then Exit;
    T := CB.Tag;
    for I := 0 to PFormat.ControlCount - 1 do
    begin
        ChildControl := PFormat.Controls[I];
        if ChildControl <> nil then
        begin
            if (ChildControl.Tag = T) and (ChildControl.Name <> CB.Name) then ChildControl.Enabled := CB.Checked;
        end;
    end;

end;

procedure TEtatPCL.SwapEcranFormat(OkOk:Boolean);
begin
    PFormat.Visible := OKOK;
    Pages.Enabled := not OkOk;
    HPB.Enabled := not OkOk;
    PanelFiltre.Enabled := not OkOk;
end;

procedure TEtatPCL.InitEcranFormat(bDetail:boolean);
var
    St, St1:string;
begin
    fbFocusFormatDetail := bDetail;
    if bDetail then St := FFormatDetail.Text
    else St := FFormat.Text;
    if Pos(';', St) = 0 then
    begin
        CBMP.Checked := TRUE;
        FMP.Text := St;
        Exit;
    end;
    if St <> '' then St1 := ReadTokenSt(St);
    if St1 <> '' then
    begin
        CBMP.Checked := TRUE;
        FMP.Text := St1;
    end;
    if St <> '' then St1 := ReadTokenSt(St);
    if St1 <> '' then
    begin
        CBMN.Checked := TRUE;
        FMN.Text := St1;
    end;
    if St <> '' then St1 := ReadTokenSt(St);
    if St1 <> '' then
    begin
        CBMU.Checked := TRUE;
        FMU.Text := St1;
    end;
end;

procedure TEtatPCL.FMPChange(Sender:TObject);
var
    ED:TEdit;
    T, i:Integer;
    TL:TLabel;
    ChildControl:tControl;
begin
    ED := TEdit(Sender);
    if ED = nil then Exit;
    T := ED.Tag;
    TL := nil;
    for I := 0 to PFormat.ControlCount - 1 do
    begin
        ChildControl := PFormat.Controls[I];
        if ChildControl <> nil then
        begin
            if (ChildControl is TLabel) and (ChildControl.Tag = T) then TL := TLabel(ChildControl);
        end;
    end;
    if TL = nil then Exit;
    if ED.Text = '' then
        TL.Caption := '4871498.25'
    else
    begin
        TL.Caption := FormatFloat(ED.Text, 1481498.25);
    end;
end;

procedure TEtatPCL.BCValideClick(Sender:TObject);
var
    St:string;
begin
    St := '';
    if CBMP.Checked then
        St := St + FMP.Text
    else
        St := '#,##'+FaitMasqueDecimales+')';
    if CBMN.Checked then St := St + ';' + FMN.Text;
    if CBMU.Checked then St := St + ';' + FMU.Text;
    if fbFocusFormatDetail then FFormatDetail.Text := St
    else FFormat.Text := St;
    SwapEcranFormat(FALSE);
    FResolChange(nil);
    FResolDetailChange(nil);
    if fbFocusFormatDetail then FFormatDetail.SetFocus
    else FFormat.SetFocus;
end;

procedure TEtatPCL.BCAbandonClick(Sender:TObject);
begin
    SwapEcranFormat(FALSE);
end;

procedure TEtatPCL.FFormatKeyDown(Sender:TObject;var Key:Word;
    Shift:TShiftState);
begin
    if not FFormat.Focused then Exit;
    if Key = VK_F5 then
    begin
        SwapEcranFormat(TRUE);
        InitEcranFormat(False);
        CBMP.SetFocus;
    end;
end;

procedure TEtatPCL.FFormatDetailKeyDown(Sender:TObject;var Key:Word;
    Shift:TShiftState);
begin
    if not FFormatDetail.Focused then Exit;
    if Key = VK_F5 then
    begin
        SwapEcranFormat(TRUE);
        InitEcranFormat(True);
        CBMP.SetFocus;
    end;
end;

function TEtatPCL.GetFichierMaquette:string;
var
    CBMaquette:THCritMaskEdit;
    lStCheminMaquette : string;
    lStCrit1, lStCrit2, lStCrit3, lStNom : string;
    lStPredefini : string;
    lRetourExtract : integer;
begin
  if PBilan.TabVisible then CBMaquette := ZMAQUETTEBIL
  else CBMAQUETTE := ZMAQUETTE;
  lStNom := '';
  Result := '';
  
  // Nom du fichier
  if (ctxPCL in V_PGI.PGIContexte) then
  begin
    if Valeur(CBMaquette.Text)>0 then
      lStNom := CBMaquette.Plus + Format('%.03d', [StrToInt(CBMaquette.Text)]) + '.TXT'
    else exit;
  end
  else lStNom := CBMaquette.Text;
  // Critère n°1
  lStCrit1 :=  'ETATSYNTH';
  // Critère n°2
  case TE of
    esSIG, esSIGA : lStCrit2 := 'SIG';
    esBIL, esBILA : lStCrit2 := 'BIL';
    esCR, esCRA : lStCrit2 := 'CR';
  end;
  // Critère n°3
  if (ctxPCL in V_PGI.PGIContexte) then
    lStCrit3 := CBMAQUETTE.Text
  else
  begin
    // Nous sommes hélas obligés de passer crit3 en paramètre à la fonction.
    // Ce n'est pas génial mais faute de mieux ...
    lStCrit3 := GetColonneSQL('YFILESTD','YFS_CRIT3','YFS_CODEPRODUIT="COMPTA" AND YFS_NOM="'+lStNom
                              +'" AND YFS_CRIT1="ETATSYNTH" AND YFS_CRIT2="'
                              +lStCrit2+'"');
  end;
  // Suppression du fichier sur disque pour récupération systématique
  lStPredefini := GetColonneSQL('YFILESTD','YFS_PREDEFINI','YFS_CODEPRODUIT="COMPTA" AND YFS_NOM="'+lStNom
                              +'" AND YFS_CRIT1="ETATSYNTH" AND YFS_CRIT2="'
                              +lStCrit2+'"');
  lStCheminMaquette := AGL_YFILESTD_GET_PATH ('COMPTA', lStNom, lStCrit1, lStCrit2, lStCrit3, '', '', V_PGI.LanguePrinc, lStPredefini);
  DeleteFile (lStCheminMaquette);
  // Récupération du fichier en base
  lRetourExtract := AGL_YFILESTD_EXTRACT( lStCheminMaquette,
                                        'COMPTA',lStNom,lStCrit1,lStCrit2,lStCrit3,'','',False,V_PGI.LanguePrinc,'STD');
  // Si rien au niveau Cabinet, on regarde au niveau Cegid
  if lRetourExtract <> - 1 then
    lRetourExtract := AGL_YFILESTD_EXTRACT( lStCheminMaquette,
                                        'COMPTA',lStNom,lStCrit1,lStCrit2,lStCrit3);

  if (lRetourExtract = -1) then Result := lStCheminMaquette
  else
  begin
    if lRetourExtract <> -1 then
    begin
      PGIInfo(AGL_YFILESTD_GET_ERR(lRetourExtract) + #13#10 + lStCheminMaquette);
    end;
    Result := ''; // Erreur, fichier non trouvé
  end;
end;

procedure TEtatPCL.OnClickAvecSituation(Sender:TObject);
var
    Indice:integer;
    AvecSitu:TCheckBox;
    HSIT:THCritMaskEdit;
    FD1, FD2:TMaskEdit;
    HV : THValComboBox;
begin
    Indice := StrToInt(Copy(TCheckBox(Sender).Name, Length(TCheckBox(Sender).Name), 1));
    AvecSitu := TCheckBox(FindComponent('AvecSitu' + IntToStr(Indice)));
    if AvecSitu <> nil then
    begin
        HSIT := THCritMaskEdit(FindComponent('Histobal' + IntToStr(Indice)));
        HSIT.Enabled := AvecSitu.Checked;
        if not HSIT.Enabled then
        begin
            HSIT.Text := '';
            HSIT.Hint := '';
        end;
        FD1 := TMaskEdit(FindComponent('FD' + IntToStr(Indice) + '1'));
        if FD1 = nil then Exit;
        FD2 := TMaskEdit(FindComponent('FD' + IntToStr(Indice) + '2'));
        if FD2 = nil then Exit;
        FD1.Enabled := not AvecSitu.Checked;
        FD2.Enabled := not AvecSitu.Checked;
        if not AvecSitu.Checked then
        begin
          HV := THvalComboBox(FindComponent('Exo' + IntToStr(Indice)));
          if HV <> nil then
            if (HV.Value='') then HV.Value := CExerciceVersRelatif(VH^.Encours.Code);
        end;
    end;
end;

procedure TEtatPCL.EnableCritCol(Col:integer);
var
    F:TEdit;
    Tp:THValComboBox;
    CB:TCheckBox;
    HIST:THCritMaskEdit;
    OkCol:TCheckBox;
begin
    F := TEdit(FindComponent('Formule' + IntToStr(Col)));
    Tp := THValComboBox(FindComponent('Type' + IntToStr(Col)));
    CB := TCheckBox(FindComponent('AvecSitu' + IntToStr(Col)));
    HIST := THCritMaskEdit(FindComponent('HistoBal' + IntToStr(Col)));
    OkCol := TCheckBox(FindComponent('OkCol' + IntToStr(Col)));
    if ((F <> nil) and (Tp <> nil)) then
    begin
    // F.Enabled := Tp.Value <> 'RUB';
        F.Enabled := Tp.Value = 'FOR'; // CA - 28/12/2001 - cas ou TP.Value indéfini
        if (CB <> nil) then
            CB.Enabled := (Tp.Value = 'RUB') and (OkCol.Checked = True);
        if HIST <> nil then HIST.Enabled := CB.Checked;
    end;
end;

procedure TEtatPCL.OnBalSitClick(Sender:TObject);
var
    Exo:string;
    HM:THCritMaskEdit;
    HV:THValComboBox;
    FD1, FD2:TMaskEdit;
    Indice:integer;
    Q:TQuery;
    GB:TGroupBox;
begin
    Exo := '';
    Indice := 0;
    GB := nil;
    HM := THCritMaskEdit(Sender);
    if HM = nil then exit;
    if PBilan.TabVisible then
    begin
        if HM.Name = 'BILBALSIT' then HV := BILEXO
        else HV := BILEXOCOMP;
    end
    else
    begin
        Indice := StrToInt(Copy(HM.Name, Length(HM.Name), 1));
        HV := THvalComboBox(FindComponent('Exo' + IntToStr(Indice)));
    end;
    if HV <> nil then
    begin
      Exo := CRelatifVersExercice(HV.Value);
    end;
(*    FQ 13023 - CA - 11/08/2004 : choix de la balance de situation indépendante de l'exercice déjà sélectionné
    if VH^.bAnalytique then LookUpBalSit(THCritMaskEdit(Sender), 'G/A', Exo, '')
    else LookUpBalSit(THCritMaskEdit(Sender), 'GEN', Exo, '');
    *)
    if VH^.bAnalytique then LookUpBalSit(THCritMaskEdit(Sender), 'G/A', '', '')
    else LookUpBalSit(THCritMaskEdit(Sender), 'GEN', '', '');
  // Mise à jour des infos suivant le choix de la balance
    Q := OpenSQL('SELECT BSI_LIBELLE,BSI_DATE1,BSI_DATE2,BSI_EXERCICE FROM CBALSIT WHERE BSI_CODEBAL="' + THCritMaskEdit(Sender).Text + '"', True);
    if not Q.Eof then
    begin
    // 1 - Mise à jour du hint avec le libellé
        HM.Hint := Q.FindField('BSI_LIBELLE').AsString;
    // 2 - Mise à jour des dates
        if PBilan.TabVisible then
        begin
            if HM.Name = 'BILBALSIT' then
            begin
                FD1 := BILDATE;
                FD2 := BILDATE_;
            end
            else
            begin
                FD1 := BILDATECOMP;
                FD2 := BILDATECOMP_;
            end;
        end
        else
        begin
            FD1 := TMaskEdit(FindComponent('FD' + IntToStr(Indice) + '1'));
            if FD1 = nil then Exit;
            FD2 := TMaskEdit(FindComponent('FD' + IntToStr(Indice) + '2'));
            if FD2 = nil then Exit;
            GB := TGroupBox(FindComponent('GB' + IntToStr(Indice)));
            if GB = nil then Exit;
        end;
        FD1.Text := Q.FindField('BSI_DATE1').AsString;
        FD2.Text := Q.FindField('BSI_DATE2').AsString;
        if not PBilan.TabVisible then
            if GB <> nil then
                GB.Tag := 5;
        HV.Value := CExerciceVersRelatif(Q.FindField('BSI_EXERCICE').AsString);
        if not PBilan.TabVisible then
            if GB <> nil then
                GB.Tag := Indice;
    // 3 - et de l'analytique
        if VH^.bAnalytique then
        begin
        end;
    end;
    Ferme(Q);
end;

procedure TEtatPCL.ZMAQUETTEChange(Sender:TObject);
begin
    THCritMaskEdit(Sender).UpdateLibelle;
end;

procedure TEtatPCL.FAfficheClick(Sender:TObject);
var
    i:integer;
    AvecSitu:TCheckBox;
    HSIT:THCritMaskEdit;
    Typ:THValComboBox;
    OkCol:TCheckBox;
begin
    for i := 0 to 3 do
    begin
        AvecSitu := TCheckBox(FindComponent('AvecSitu' + IntToStr(i)));
        if AvecSitu <> nil then
        begin
            Typ := THValComboBox(FindComponent('Type' + IntToStr(i)));
            OkCol := TCheckBox(FindComponent('OkCol' + IntToStr(i)));
            AvecSitu.Enabled := (Typ.Value = 'RUB') and (OkCol.Checked) and (not
                VH^.BAnalytique);
            if not AvecSitu.Enabled then
                AvecSitu.Checked := False;
            HSIT := THCritMaskEdit(FindComponent('Histobal' + IntToStr(i)));
            if HSIT <> nil then
            begin
                HSIT.Enabled := AvecSitu.Checked;
                if not HSIT.Enabled then HSIT.Text := '';
            end;
        end;
    end;
end;

procedure TEtatPCL.BAideClick(Sender:TObject);
begin
    CallHelpTopic(Self);
end;

procedure TEtatPCL.FNatureCptChange(Sender:TObject);
var lStAxe : string;
begin
    if not gbChargementFiltre then
    begin
      FSECTION.Clear;
      FSECTION_.Clear;
    end;

    // GCO - 17/07/2006 - FQ 16318
    lStAxe := Copy( FNatureCpt.Value, 2, 1);
    if lStAxe = '1' then
      FSECTION.DataType := 'TZSECTION'
    else
      FSECTION.DataType := 'TZSECTION' + lStAxe;

    FSECTION_.DataType := FSECTION.DataType;
    // FIN - GCO
end;

procedure TEtatPCL.edt_SectionDblClick(Sender:TObject);
var
    Lefb:TFichierBase;
    stCodeDe, stCodeA, stStock, stTemp, stTexte:string;
    i, j:Integer;
    edt_Text:TEdit;
begin
    Lefb := fbAxe1;
    stCodeDe := '';
    stCodeA := '';
    StStock := '';
    edt_Text := TEdit(Sender);
    if edt_Text = nil then exit;
    stTexte := edt_Text.Text;

    if stTexte <> '' then
    begin
        i := Pos(':', stTexte);
        if i > 0 then
        begin
            stCodeDe := Copy(stTexte, 1, i - 1);
            stCodeA := Copy(stTexte, i + 1, Length(stCodeDe));
            StStock := Copy(stTexte, 2 * Length(stCodeDe) + 2, Length(stTexte));

            stCodeDe := FindEtReplace(stCodeDe, ',', ';', True);
            stCodeA := FindEtReplace(stCodeA, ',', ';', True);
        end
        else
        begin
            i := Pos('&', stTexte);
            j := Pos('|', stTexte);
            if (i = 0) and (j = 0) then stCodeDe := stTexte;
            if (i <> 0) and (j = 0) then
            begin
                stCodeDe := Copy(stTexte, 1, i - 2);
                StStock := Copy(stTexte, i, Length(stTexte));
            end;
            if (j <> 0) and (i = 0) then
            begin
                stCodeDe := Copy(stTexte, 1, j - 2);
                StStock := Copy(stTexte, j, Length(stTexte));
            end;
            if (i <> 0) and (j <> 0) then
            begin
                stCodeDe := Copy(stTexte, 1, Min(i, j) - 2);
                StStock := Copy(stTexte, Min(i, j), Length(stTexte));
            end;
            stCodeDe := FindEtReplace(stCodeDe, ',', ';', True);
            stCodeA := stCodeDe;
        end;
    end;
    if StStock <> '' then
        if StStock[1] = ',' then StStock := Copy(StStock, 2, Length(StStock));
    ChoixTableLibrePourRub(Lefb, FNatureCpt.Value, StStock, stCodeDe, stCodeA);

    stCodeDe := FindEtReplace(stCodeDe, ';', ',', True);
    stCodeA := FindEtReplace(stCodeA, ';', ',', True);

    if stCodeA <> stCodeDe then StTemp := stCodeDe + ':' + stCodeA
    else StTemp := stCodeDe;
    if StStock <> '' then
    begin
        if StStock[1] <> ',' then StStock := ',' + StStock;
        StTemp := StTemp + StStock;
    end;
    edt_Text.Text := StTemp;
end;

procedure TEtatPCL.edt_SectionKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
var
    edt_Text:TEdit;
begin
    if Key <> VK_DELETe then exit;
    edt_Text := TEdit(Sender);
    if edt_Text = nil then exit else edt_Text.Text := '';
end;

procedure TEtatPCL.rdg_AnalytiqueClick(Sender:TObject);
begin
    if rdg_Analytique.Value = 'TLI' then
    begin
        pnl_TLibres.Visible := False;
        pnl_Comptes.Visible := True;
    end
    else
    begin
        pnl_TLibres.Visible := True;
        pnl_Comptes.Visible := False;
    end;
end;

procedure TEtatPCL.FSECTIONKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = '=' then Key := #0;
end;

procedure TEtatPCL.BILEXOChange(Sender:TObject);
begin
    if THValComboBox(Sender).Name = 'BILEXO' then
    begin
        if BILBALSIT.Text <> '' then exit;
        CExoRelatifToDates(BILEXO.Value, BILDATE, BILDATE_);
    // Si changement d'exercice et balance de situation sélectionnée, on 'vide' la balance
        BILBALSIT.Text := '';
        BILBALSIT.Hint := '';
    end
    else if THValComboBox(Sender).Name = 'BILEXOCOMP' then
    begin
        if BILBALSITCOMP.Text <> '' then exit;
        CExoRelatifToDates(BILEXOCOMP.Value, BILDATECOMP, BILDATECOMP_);
    // Si changement d'exercice et balance de situation sélectionnée, on 'vide' la balance
        BILBALSITCOMP.Text := '';
        BILBALSITCOMP.Hint := '';
    end;
end;

procedure TEtatPCL.BILAVECSITClick(Sender:TObject);
var
    AvecSitu:TCheckBox;
    BalSit:THCritMaskEdit;
begin
    if TCheckBox(Sender).Name = 'BILAVECSIT' then
    begin
        AvecSitu := BILAVECSIT;
        BalSit := BILBALSIT;
    end
    else
    begin
        AvecSitu := BILAVECSITCOMP;
        BalSit := BILBALSITCOMP;
    end;
    BalSit.Enabled := AvecSitu.Checked;
    if not BalSit.Enabled then
    begin
        BalSit.Text := '';
        BalSit.Hint := '';
    end;
    BILDATE.Enabled := not BILAVECSIT.Checked;
    BILDATE_.Enabled := not BILAVECSIT.Checked;
    BILDATECOMP.Enabled := not BILAVECSIT.Checked;
    BILDATECOMP_.Enabled := not BILAVECSIT.Checked;
end;

procedure TEtatPCL.BILCOMPClick(Sender:TObject);
begin
    GBBILANCOMP.Enabled := BILCOMP.Checked;
    if not BILCOMP.Checked then
    begin
        BILAVECSITCOMP.Checked := False;
        BILBALSITCOMP.Text := '';
    end;
end;

function TEtatPCL.RecupCritBilan:Boolean;
var
    i:Integer;
    CodeExo:string;
begin
  // Colonne 1 & 2
    for i := 1 to 2 do
        with Crit do
        begin
            Col[i - 1].Actif := True;
            CodeExo := CRelatifVersExercice(BILEXO.Value);
            QuelDateDeExo(CodeExo, Col[i - 1].Exo);
            Col[i - 1].Date1 := StrToDate(BILDATE.Text);
            Col[i - 1].Date2 := StrToDate(BILDATE_.Text);
            Col[i - 1].BalSit := BILBALSIT.Text;
            Col[i - 1].IsFormule := FALSE;
        end;
  // Colonne 3
    with Crit do
    begin
        Col[2].Actif := True;
        Col[2].StFormule := '=[COL1]-[COL2]';
        Col[2].IsFormule := TRUE;
        Col[2].StTitre := 'Simple : ' + #13 + #10 + 'Variation en valeur';
    end;
  // Colonne 4
    Crit.Col[3].Actif := BILCOMP.Checked;
    if BILCOMP.Checked then
    begin
        with Crit do
        begin
            CodeExo := CRelatifVersExercice(BILEXOCOMP.Value);
            QuelDateDeExo(CodeExo, Col[3].Exo);
            Col[3].Date1 := StrToDate(BILDATECOMP.Text);
            Col[3].Date2 := StrToDate(BILDATECOMP_.Text);
            Col[3].BalSit := BILBALSITCOMP.Text;
            Col[3].IsFormule := FALSE;
        end;
    end;
    Result := TRUE;
end;

procedure TEtatPCL.Timer1Timer(Sender:TObject);
begin
    Timer1.Enabled := False;

  // Chargement des filtres de l'état dans la combo FFiltres
    if FCritEdtChaine.FiltreUtilise <> '' then
    begin
    // Chargement du Filtre utilisé par les états chainés
        FFiltres.ItemIndex := FFiltres.Items.Indexof(FCritEdtChaine.FiltreUtilise);
        FFiltres.OnChange(FFiltres);
    end;

    BValider.click;
    Close;
end;

procedure TEtatPCL.InitEtatChaine;
begin
  // Positionnement du itemindex de la combo avec le nom du filtre
    FFiltres.ItemIndex := FFiltres.Items.Indexof(FCritEdtChaine.FiltreUtilise);
    LoadFiltre(FNomFiltre, FFiltres, Pages);

  // Si pas de filtre valable, on met des paramètres d'édition par défaut
    if (TE = esBIL) or (TE = esBILA) then
    begin
        if (BILEXO.Value = '') or (ZMAQUETTEBIL.Text = '') then
        begin
            ZMAQUETTEBIL.Text := '7';
            BILEXO.Value := CExerciceVersRelatif(VH^.Encours.Code);
            BILDATE.Text := DateToStr(VH^.Encours.Deb);
            BILDATE_.Text := DateToStr(VH^.Encours.Fin);
        end;
    end
    else
    begin
        if (not OkCol1.Checked) or (ZMAQUETTE.Text = '') then
        begin
            ZMAQUETTE.Text := '7';
            OkCol1.Checked := True;
            Type1.Value := 'RUB';
            Exo1.Value := CExerciceVersRelatif(VH^.Encours.Code);
            FD11.Text := DateToStr(VH^.Encours.Deb);
            FD12.Text := DateToStr(VH^.Encours.Fin);
        end;
    end;
    
    // GCO - 29/11/2004 - FQ13484
    if FCritEdtChaine.UtiliseCritStd then
    begin
      // Chargement des critères standards des états chainés
      if (TE = esBIL) or (TE = esBILA) then
      begin
        BILEXO.Value := FCritEdtChaine.Exercice.Code;
        BILDATE.Text  := DateToStr(FCritEdtChaine.Exercice.Deb);
        BILDATE_.Text  := DateToStr(FCritEdtChaine.Exercice.Fin);
      end else
      begin
        Exo1.Value := FCritEdtChaine.Exercice.Code;
        FD11.Text  := DateToStr(FCritEdtChaine.Exercice.Deb);
        FD12.Text  := DateToStr(FCritEdtChaine.Exercice.Fin);
      end;
    end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2003
Modifié le ... :   /  /
Description .. : Ajout d'un filtre
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.InitAddFiltre(T:TOB);
var
    Lines: HTStrings;
begin
    Lines := HTStringList.Create;
    SauveCritMemoire(Lines, Pages);
    FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
    Lines.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.InitGetFiltre(T:TOB);
var
    Lines: HTStrings;
begin
    Lines := HTStringList.Create;
    SauveCritMemoire(Lines, Pages);
    FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
    Lines.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2003
Modifié le ... :   /  /
Description .. : Chargement et sélection du filtre sélectionné
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.InitSelectFiltre(T:TOB);
var
    Lines: HTStrings;
    i:integer;
    stChamp, stVal:string;
begin
    if T = nil then exit;
    Lines := HTStringList.Create;
    for i := 0 to T.Detail.Count - 1 do
    begin
        stChamp := T.Detail[i].GetValue('N');
        stVal := T.Detail[i].GetValue('V');
        Lines.Add(stChamp + ';' + stVal);
    end;
    VideFiltre(FFiltres, Pages, False);
    ChargeCritMemoire(Lines, Pages);
    Lines.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2003
Modifié le ... :   /  /
Description .. : Conversion des anciens filtres au format XML
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.ParseParamsFiltre(Params: HTStrings);
var
    T:TOB;
begin
    FListeByUser.AddVersion;
    T := FListeByUser.Add;
  //en position 0 de Params se trouve le nom du filtre
    T.PutValue('NAME', XMLDecodeSt(Params[0]));
    T.PutValue('USER', '---');
    Params.Delete(0);
    FListeByUser.AffecteTOBFiltreMemoire(T, Params);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/12/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.UpgradeFiltre(T:TOB);
begin
  //
end;
procedure TEtatPCL.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F10 :
      begin
        BValiderClick(nil);
        Key := 0;
      end;
  end;
end;

procedure TEtatPCL.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then CloseInsidePanel(Self) ;
end;

procedure TEtatPCL.LanceEdition ( pbExport : boolean );
var
    SFormule, SFormat, SFormatDetail:array of string;
    i:integer;
begin
(*
    if not (ctxPCL in V_PGI.PGIContexte) then
    begin
        if (Crit.TE = esBil) or (Crit.TE = esBILA) then
        begin
            if not FileExists(ZMAQUETTEBIL.Text) then
            begin
                MsgBox.Execute(10, Caption, '');
                exit;
            end;
        end
        else
        begin
            if not FileExists(ZMAQUETTE.Text) then
            begin
                MsgBox.Execute(10, Caption, '');
                exit;
            end;
        end;
    end;
    *)
    if not RecupCrit then Exit;
    SetLength(SFormule, Crit.NbColActif);
    SetLength(SFormat, Crit.NbColActif);
    SetLength(SFormatDetail, Crit.NbColActif);
    for i := 0 to Crit.NbColActif - 1 do if Crit.Col[i].Actif then
        begin
            SFormule[i] := Crit.Col[i].StFormule;
            SFormat[i] := Crit.Col[i].StFormat;
            SFormatDetail[i] := Crit.Col[i].StFormatDetail;
        end;
    Crit.bExport := pbExport;
    LanceLiasse(fFichierMaquette, FALSE, FVoirNom.Checked, sFormule, sFormat,
            sFormatDetail, Crit, FALSE, FCritEdtChaine);
    SFormule := nil;
    SFormat := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/09/2005
Modifié le ... :   /  /
Description .. : Lance l'export vers Excel
Mots clefs ... :
*****************************************************************}
procedure TEtatPCL.BExportClick(Sender: TObject);
begin
  if ExJaiLeDroitConcept(ccExportListe, True) then LanceEdition ( True );
end;

procedure TEtatPCL.FSECTIONChange(Sender: TObject);
var
    AvecJoker:Boolean;
begin
    AvecJoker := EstJoker ( FSECTION.Text );
    TFaG.Visible := not AvecJoker;        { 'à' }
    TFGen.Visible := not AvecJoker;       { 'Sections de ' }
    TFGenJoker.Visible := AvecJoker;      { 'Sections' }
    FSECTION_.Visible := not AvecJoker;   { Borne supérieure Section }
    FUnEtatParSection.Enabled := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/03/2006
Modifié le ... :   /  /
Description .. : Contrôle de cohérence des critères d'édition
Mots clefs ... :
*****************************************************************}
function TEtatPCL.ControleCritere ( var vStMsg : string ) : boolean;
var
  i, iCol : integer;
  OkCol : TCheckBox;
  AvecSitu : TCheckBox;
  OkColPrec : boolean;
  HIV : THValComboBox;
  FD1,FD2 : TMaskEdit;
  Exo : TExoDate;
  iLaDate : integer;
begin
  vStMsg := '';
  Result := True;
  iLaDate := 0;
  HIV := nil;
  FD1 := nil; FD2 := nil;
  if PBilan.TabVisible then
  begin
    if (IsValidDate(BILDATE.Text) and IsValidDate(BILDATE_.Text)) then
    begin
      if not BILAVECSIT.Checked then
      begin
        Result :=True;
        iLaDate := 1;
        FillChar( Exo, SizeOf(Exo), #0);
        CQuelExercice( StrToDate(BILDATE.Text) , Exo );
        if (CRelatifVersExercice(BILEXO.Value)<>Exo.Code) then Result := False
        else
        begin
          iLaDate := 2;
          FillChar( Exo, SizeOf(Exo), #0);
          CQuelExercice( StrToDate(BILDATE_.Text) , Exo );
          if (CRelatifVersExercice(BILEXO.Value)<>Exo.Code) then Result := False;
        end;
        if not Result then
        begin
          vStMsg := 'Problème au niveau du paramétrage des dates';
          PGIBox ( vStMsg );
          QuelDateDeExo(CRelatifVersExercice(BILEXO.Value),Exo) ;
          if (iLaDate = 1) then BILDATE.Text := DateToStr(Exo.Deb) else BILDATE_.Text := DateToStr(Exo.Fin);
          exit;
        end;
      end;
      if ((BILCOMP.Checked) and (not BILAVECSITCOMP.Checked)) then
      begin
        if (IsValidDate(BILDATECOMP.Text) and IsValidDate(BILDATECOMP_.Text)) then
        begin
          Result :=True;
          iLaDate := 1;
          FillChar( Exo, SizeOf(Exo), #0);
          CQuelExercice( StrToDate(BILDATECOMP.Text) , Exo );
          if (CRelatifVersExercice(BILEXOCOMP.Value)<>Exo.Code) then Result := False
          else
          begin
            iLaDate := 2;
            FillChar( Exo, SizeOf(Exo), #0);
            CQuelExercice( StrToDate(BILDATECOMP_.Text) , Exo );
            if (CRelatifVersExercice(BILEXOCOMP.Value)<>Exo.Code) then Result := False;
          end;
          if not Result then
          begin
            vStMsg := 'Problème de date sur le comparatif';
            PGIBox ( vStMsg );
            QuelDateDeExo(CRelatifVersExercice(BILEXOCOMP.Value),Exo) ;
            if (iLaDate = 1) then BILDATECOMP.Text := DateToStr(Exo.Deb) else BILDATECOMP_.Text := DateToStr(Exo.Fin);
            exit;
          end;
        end;
      end;
    end;
  end else
  begin
    iCol := 0;
    okColPrec := True;
    try
      for i := 1 to 4 do
      begin
        OkCol := TCheckBox(FindComponent('OkCol' + IntToStr(i)));
        if (assigned(OkCol) and OkCol.Checked) then
        begin
          if (not OkColPrec) then
          begin
            iCol := i+4;
            break;
          end;
          AvecSitu := TCheckBox(FindComponent('AvecSitu' + IntToStr(i)));
          if (Assigned(AvecSitu) and (AvecSitu.Checked)) then continue;
          HIV := THValComboBox(FindComponent('Type' + IntToStr(i)));
          if (Assigned(HIV) and (HIV.Value = 'RUB')) then
          begin
            HIV := THValComboBox(FindComponent('Exo' + IntToStr(i)));
            if (Assigned(HIV) and (HIV.Value <> '')) then
            begin
              FD1 := TMaskEdit(FindComponent('FD' + IntToStr(i) + '1'));
              FD2 := TMaskEdit(FindComponent('FD' + IntToStr(i) + '2'));
              if (IsValidDate(FD1.Text) and IsValidDate(FD2.Text)) then
              begin
                iLaDate := 1;
                FillChar( Exo, SizeOf(Exo), #0);
                (* On vérifie désormais que les dates choisies appartiennent à l'exercice *)
                CQuelExercice( StrToDate(FD1.Text) , Exo );
                if (Exo.Code <> CRelatifVersExercice ( HIV.Value )) then iCol := i else
                begin
                  iLaDate := 2;
                  FillChar( Exo, SizeOf(Exo), #0);
                  CQuelExercice( StrToDate(FD2.Text) , Exo );
                  if (Exo.Code <> CRelatifVersExercice ( HIV.Value )) then iCol := i
                  else iCol := 0;
                end;
                (* Ancienne méthode ...
                if (CQuelExercice( StrToDate(FD1.Text) , Exo )
                  and CQuelExercice( StrToDate(FD2.Text) , Exo )) then iCol := 0
                else iCol := i;
                *)
                OkColPrec := True;
              end else iCol := i;
            end else iCol := i;
          end;
        end else OkColPrec := False;
        if (iCol<>0)  then break;
      end;
    finally
      if iCol = 0 then vStMsg := ''
      else if (iCol < 5) then
      begin
        vStMsg := 'Problème au niveau du paramétrage des dates de la colonne '+IntToStr(iCol);
        PGIBox (vStMsg);
        if assigned(HIV) then QuelDateDeExo(CRelatifVersExercice(HIV.Value),Exo) ;
        if ((iLaDate = 1) and assigned (FD1)) then FD1.Text := DateToStr(Exo.Deb) else
          if assigned(FD2) then FD2.Text := DateToStr(Exo.Fin);
      end
      else
      begin
        vStMsg := 'La colonne '+IntToStr(iCol-1-4)+' doit être sélectionnée';
        PGIBox (vStMsg);
      end;
      Result := (iCol=0);
    end;
  end;
end;

procedure TEtatPCL.InitComboPlanCorrespondance;
var
  bOkAdd : boolean;
  ValDefaut : string;
begin
  bOkAdd := False;
  if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) then
  begin
    if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='1' then ValDefaut := '1'
    else if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='2' then ValDefaut := '2'
    else ValDefaut := '';
  end else ValDefaut := '';

  if GetParamSocSecur('SO_CORSGE1', True, True) then bOkAdd := True;
  if GetParamSocSecur('SO_CORSGE2', True, True) then bOkAdd := True;
  if not bOkAdd then
  begin
    FCORRESPONDANCE.Value := '';
    FCORRESPONDANCE.Enabled := False;
  end else FCORRESPONDANCE.Value := ValDefaut;
end;

(*
procedure TEtatPCL.InitRadioGroupDevise;
begin
  { Si plusieurs devises gérées dans le dossier, alors on n'affiche que la devise de tenue }
  // à faire ...
end;
*)
end.
