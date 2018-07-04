unit MulSuiviMP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB, DBTables, Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, UTOB, UtilPGI,FileCtrl,ParamSoc,
  HRichOLE, MulSMPUtil, HStatus, TofVerifRib
  ,uLibWindows // pour CVireLigne
{$IFNDEF IMP}
  ,SaisBor
{$IFNDEF CCMP}
  , HRichOLE
{$ENDIF}
{$ENDIF}
  ;

Procedure SelectSuiviMP ( smp : TSuiviMP ) ;

type
  TFMulSuiviMP = class(TFMul)
    HM: THMsgBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    TE_DEBIT: TLabel;
    TE_DEBIT_: TLabel;
    TE_CREDIT: TLabel;
    TE_CREDIT_: TLabel;
    TE_ETABLISSEMENT: THLabel;
    E_DEBIT: THCritMaskEdit;
    E_DEBIT_: THCritMaskEdit;
    E_CREDIT: THCritMaskEdit;
    E_CREDIT_: THCritMaskEdit;
    E_ETABLISSEMENT: THValComboBox;
    PEcritures: TTabSheet;
    Bevel6: TBevel;
    TE_JOURNAL: THLabel;
    TE_NATUREPIECE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    TE_EXERCICE: THLabel;
    TE_DATEECHEANCE: THLabel;
    TE_DATEECHEANCE2: THLabel;
    E_EXERCICE: THValComboBox;
    E_JOURNAL: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHEREAN: TEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREENC: TEdit;
    XX_WHEREVIDE: TEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel17: THLabel;
    E_DEVISE: THValComboBox;
    HLabel7: THLabel;
    CATEGORIE: THValComboBox;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    HDiv: THMsgBox;
    XX_WHEREMP: TEdit;
    HLabel1: THLabel;
    E_NUMENCADECA: THCritMaskEdit;
    PBanqueLot: TTabSheet;
    Bevel7: TBevel;
    HLabel3: THLabel;
    XNOMLOT1: THCritMaskEdit;
    HLabel5: THLabel;
    XNOMLOT2: THCritMaskEdit;
    HLabel6: THLabel;
    E_BANQUEPREVI: THCritMaskEdit;
    HLabel8: THLabel;
    E_BANQUEPREVI_: THCritMaskEdit;
    XX_WHERELOT: TEdit;
    BCtrlRib: TToolbarButton97;
    FE_NUMTRAITECHQ: TLabel;
    FE_NUMTRAITECHQ_: TLabel;
    E_NUMTRAITECHQ_: THCritMaskEdit;
    E_NUMTRAITECHQ: THCritMaskEdit;
    XX_WHERETRACHQ: TEdit;
    BRecapTraite: TToolbarButton97;
    XX_WHERENATCPT: TEdit;
    FTID: TCheckBox;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    cFactCredit: TCheckBox;
    TraiteEdite: TCheckBox;
    XX_WHEREPROFIL: TEdit;
    XX_WHEREMONTANT: TEdit;
    CRechVide: TCheckBox;
    BRAZLot: TToolbarButton97;
    NewFiltreMP: TToolbarButton97;
    BKillLot: TToolbarButton97;
    BSwapSelect: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure CATEGORIEChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BCtrlRibClick(Sender: TObject);
    procedure FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
    procedure E_NUMTRAITECHQChange(Sender: TObject);
    procedure cFactCreditClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override;
    procedure FTIDClick(Sender: TObject);
    procedure TraiteEditeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//    procedure FListeEnter(Sender: TObject);
    procedure XNOMLOT1Change(Sender: TObject);
    procedure CRechVideClick(Sender: TObject);
    procedure BRAZLotClick(Sender: TObject);
    procedure NewFiltreMPClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure BKillLotClick(Sender: TObject);
//    procedure DessineCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);  // BPY 29/01/2003 Correctin de la fiche 11794
    procedure FListeFlipSelection(Sender: TObject);
    procedure BSwapSelectClick(Sender: TObject);
  private    { Déclarations privées }
    TobMvt : Tob ;
    CodeLot : String ;
    SwapSelect : Boolean ;						 // Gestion du mode selection inversé
    gszCaption : String;               // Pour les messages
    procedure InitCriteres ;
    procedure InitE_NUMTRACHQ(Invisible : Boolean) ;
    procedure InitContexte ;
    procedure ReinitWhereNatCpt(RAZ : Boolean) ;
    procedure InitEnc ;
    procedure InitDec ;
    procedure PreciseMP ;
    Function  ConstitueTobMvt : Boolean ;
    procedure VireInutiles ;
    procedure ToutMarquer ;
    procedure MarqueOrigine ;
    procedure PrechargeOrigines ;
    procedure ValideLot ;
    procedure ClickModifRib ;
    Function  GereXNomLot : Boolean ;
    Function  CalculSoldeSelection : Double ;
    Procedure AfficheSoldeSelection ;
//    Procedure InitTitreContexte;
 public
    smp             : TSuiviMP ;
    SorteLettre : TSorteLettre ;
  end;

implementation

uses FichComm, SaisLot, Filtre, DetruitLot ;

{$R *.DFM}

Var SmpRes : String = '' ;

Procedure SelectSuiviMP ( smp : TSuiviMP ) ;
var
    X : TFMulSuiviMP ;
    PP : THPanel ;
//    ind : tProfilTraitement ;
//    OkOk : Boolean ;
begin
    (*
    If Not ProfilOk(smp) Then
    BEGIN
        HShowMessage('0;Encaissements/Décaissements;Vous n''avez défini de profil !;E;O;O;O;','','') ;
        HShowMessage('1;Encaissements/Décaissements;Menu autres traitements / Profils utilisateur!;E;O;O;O;','','') ; Exit ;
    END ;
    *)
    (*
    Ind:=QuelProfil(smp,OkOk) ;
    If Not MonProfilOk(Ind) Then
    BEGIN
        *)
        (* GP le 18/12/2001
        smpRes:='SMPP'+IntToStr(Ord(smp)) ;
        if Blocage([smpRes],True,smpRes) then Exit ;
        *)
        (*
    END ;
    *)
    PP:=FindInsidePanel ;

    X:=TFMulSuiviMP.Create(Application) ;

    X.Q.Manuel:=True ;
    X.smp:=smp ;
    X.SorteLettre:=AttribSL(smp) ;
    X.FNomFiltre:=AttribFiltre(smp,'SV') ;
    X.Q.Liste:=AttribListe(smp) ;
    X.HelpContext:=AttribHelp(smp,TRUE) ;

    if PP=Nil then
    BEGIN
        try
            X.ShowModal ;
        finally
            X.Free ;
            (*If Not MonProfilOk(Ind) Then*)
            (* GP le 18/12/2001 Bloqueur(smpRes,False) ;*)
        end;
        Screen.Cursor:=SyncrDefault ;
    END else
    BEGIN
        InitInside(X,PP) ;
        X.Show ;
    END ;
END ;

(*========================= METHODES DE LA FORM ==============================*)
procedure TFMulSuiviMP.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;
TOBMvt:=TOB.Create('',Nil,-1) ;
// Gestion du mode "selection inversé", non opérationnel par défaut
SwapSelect := False ;
end;

procedure TFMulSuiviMP.FormShow(Sender: TObject);
begin
InitCriteres ;

    if (smp=smpDecVirEdt)   then HelpContext:=999999433;
    if (smp=smpDecVirInEdt) then HelpContext:=999999438;
    if (smp=smpDecVirBqe)   then HelpContext:=999999436;
    if (smp=smpDecVirInBqe) then HelpContext:=999999441;

if (smp<>smpEncTraEdt) And (smp<>smpEncTraEdtNC) And
   (smp<>smpDecBorEdt) And (smp<>smpDecBorEdtNC) Then
   BEGIN
   TraiteEdite.State:=cbGrayed ;
   E_NUMTRAITECHQ_.Visible:=FALSE ;
   FE_NUMTRAITECHQ_.Enabled:=FALSE ;
   E_NUMTRAITECHQ.Enabled:=FALSE ;
   FE_NUMTRAITECHQ.Enabled:=FALSE ;
   E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
   XX_WHERETRACHQ.Text:='' ;
   END ;
TraiteEditeClick(Nil) ;
  inherited;
VH^.MPModifFaite:=FALSE ; Fillchar(VH^.MPPOP,SizeOf(VH^.MPPOP),#0) ;
InitContexte ;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHEREVIDE.Text:='' ;
cFactCreditClick(Nil) ;
end;

(*============================ INITIALISATIONS ===============================*)
procedure TFMulSuiviMP.PreciseMP ;
BEGIN
If Categorie.itemIndex=0 Then CatToMP('',E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True)
                         Else CatToMP(CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True) ;
  // On ôte les virement internationnaux dans le menu BOR // SBO fiche 10594
  if smp in [smpDecBorEdt,smpDecBorEdtNC,smpDecborDec,smpDecTraPor] then
    CVireLigne( E_MODEPAIE , 'VIN' ) ;
  // Fin SBO fiche 10594
                         
(*
if smp in [smpEncTous,smpDecTous] then Exit ;
If smp in [smpEncTraEdt,smpEncTraEdtNC] then
  BEGIN
  If Categorie.itemIndex=0 Then CatToMP('',E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True)
                           Else CatToMP(CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True) ;
  END Else
  BEGIN
  Case SorteLettre of
       tslAucun  : CatToMP(CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,SorteLettre,True) ;
       tslCheque : BEGIN E_MODEPAIE.DataType:='TTMODEPAIECHEQUE' ; E_MODEPAIE.Reload ; END ;
       tslBOR    : BEGIN E_MODEPAIE.DataType:='TTMODEPAIETRAITE' ; E_MODEPAIE.Reload ; END ;
       tslTraite : BEGIN E_MODEPAIE.DataType:='TTMODEPAIETRAITE' ; E_MODEPAIE.Reload ; END ;
  //     tslVir    : BEGIN E_MODEPAIE.DataType:='TTMODEPAIECHEQUE' ; E_MODEPAIE.Reload ; END ;
       END ;
  END ;
if E_MODEPAIE.Vide then
   BEGIN
   if E_MODEPAIE.Values.Count=2 then E_MODEPAIE.Value:=E_MODEPAIE.Values[1] ;
   END else
   BEGIN
   if E_MODEPAIE.Values.Count=1 then E_MODEPAIE.Value:=E_MODEPAIE.Values[0] ;
   END ;
*)
END ;

procedure TFMulSuiviMP.ReinitWhereNatCpt(RAZ : Boolean) ;
Var StXP,StXP2,StXN,StXN2,St : String ;
BEGIN
If RAZ Then XX_WHERENATCPT.Text:='' Else
  BEGIN
  If FTID.Checked Then
    BEGIN
    If IsEnc(SMP) Then XX_WHERENATCPT.Text:='E_AUXILIAIRE="" AND G_NATUREGENE<>"TIC"'
                  Else XX_WHERENATCPT.Text:='E_AUXILIAIRE="" AND G_NATUREGENE<>"TID"';
    END Else
    BEGIN
    If IsEnc(SMP) Then XX_WHERENATCPT.Text:='E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"FOU" AND T_NATUREAUXI<>"AUC"'
                  Else XX_WHERENATCPT.Text:='E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"CLI" AND T_NATUREAUXI<>"AUD"';
    END ;
  END ;
StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
St:='(E_DEBIT+E_CREDIT-E_COUVERTURE not between '+StXN+' AND '+StXP+')' ;
//St:=St+' And (E_DEBITEURO+E_CREDITEURO-E_COUVERTUREEURO not between '+StXN2+' AND '+StXP2+')' ;
XX_WHEREMONTANT.Text:=St ;
END ;

procedure TFMulSuiviMP.InitEnc ;
BEGIN
(*
if cFactCredit.Checked then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="FC"))'
                       else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
*)
if cFactCredit.Checked then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"))'
                       else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
E_GENERAL.DataType:='TZGENCAIS' ; E_AUXILIAIRE.DataType:='TZTTOUTDEBIT' ;
if smp=smpEncTous then XX_WHERELOT.Text:='E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ;
END ;

procedure TFMulSuiviMP.InitDec ;
BEGIN
(*
if cFactCredit.Checked then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="FF"))'
                       else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
*)
if cFactCredit.Checked
   then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"))'
   else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
E_GENERAL.DataType:='TZGDECAIS' ; E_AUXILIAIRE.DataType:='TZTTOUTCREDIT' ;
if smp=smpDecTous then XX_WHERELOT.Text:='E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ;
cFactCredit.Caption:=HDiv.Mess[15] ;
END ;

procedure TFMulSuiviMP.InitE_NUMTRACHQ(Invisible : Boolean) ;
BEGIN
E_NUMTRAITECHQ_.Visible:=Not Invisible ; FE_NUMTRAITECHQ_.Visible:=Not Invisible ;
E_NUMTRAITECHQ.Visible:=Not Invisible ; FE_NUMTRAITECHQ.Visible:=Not Invisible ;
END ;

procedure TFMulSuiviMP.InitContexte ;
BEGIN
InitE_NUMTRACHQ(smp in [smpEncPreEdt,smpEncTous,smpDecVirEdt,smpDecVirBqe,smpDecVirInEdt,smpDecVirInBqe,smpEncPreBqe]) ;
XNomLot1.Plus:=AttribPlus(smp) ; XNomLot2.Plus:=AttribPlus(smp) ;
Case smp of
   smpEncDiv    : BEGIN // Encaissement divers
                  InitEnc ; CATEGORIE.ItemIndex:=0 ; Caption:=HDiv.Mess[20] ;
                  END ;
   smpEncPreEdt : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='PRE' ; Caption:=HDiv.Mess[19] ;
                  END ;
   smpEncPreBqe : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='CB' ; Caption:=HDiv.Mess[29] ;
                  END ;
   smpEncChqPor : BEGIN // Mise en portefeuille chèques clients
                  InitEnc ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[28] ;
                  END ;
    smpEncCBPor : BEGIN // Mise en portefeuille cartes bleues clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[16] ;
                  END ;
   smpEncTraPor : BEGIN // Mise en portefeuille traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[16] ;
                  END ;
   smpEncTraEdt : BEGIN // Lettres-traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[5] ;
                  END ;
 smpEncTraEdtNC : BEGIN // Lettres-traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[26] ;
//                  XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ=""' ;
                  END ;
   smpEncTraEnc : BEGIN // Traite clients à l'encaissement
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[8] ;
                  END ;
   smpEncTraEsc : BEGIN // Traite clients à l'escompte
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[9] ;
                  END ;
   smpEncChqBqe : BEGIN // Chèque clients en banque
                  InitEnc ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[31] ;
                  END ;
    smpEncCBBqe : BEGIN // CB clients en banque
                  InitEnc ; CATEGORIE.Value:='CB' ; Caption:=HDiv.Mess[6] ;
                  END ;
   smpEncTraBqe : BEGIN // Traite clients en banque
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[6] ;
                  END ;
     smpEncTous : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='' ; Caption:=HDiv.Mess[12] ;
                  END ;
   smpDecChqEdt : BEGIN // Lettres-chèques fournisseurs
                  InitDec ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[2] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[25] ;
                  END ;
 smpDecChqEdtNC : BEGIN // Lettres-chèques fournisseurs
                  InitDec ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[2] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[25] ;
                  END ;
   smpDecVirEdt : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[18] ;
                  END ;
 smpDecVirEdtNC : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[18] ;
                  END ;
   smpDecVirBqe : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[3] ;
                  END ;
 smpDecVirInEdt : BEGIN // Virements internationales fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[34] ;
                  END ;
 smpDecVirInEdtNC : BEGIN // Virements internationales fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[34] ;
                  END ;
 smpDecVirInBqe : BEGIN // Virements internationales fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[35] ;
                  END ;
   smpDecBorEdt : BEGIN // Lettres-Bor fournisseurs
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[27] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
 smpDecBorEdtNC : BEGIN // Lettres-Bor fournisseurs
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[4] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
   smpDecborDec : BEGIN // Traites fournisseurs en encaissement
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[10] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
   smpDecBorEsc : BEGIN // Traites fournisseurs à l'escompte
                  InitDec ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[11] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
   smpDecTraBqe : BEGIN // Traites fournisseurs en banque
                  InitDec ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[7] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
   smpDecTraPor : BEGIN // Mise en portefeuille traite clients
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[17] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
      smpDecDiv : BEGIN // Décaissements divers
                  InitDec ; CATEGORIE.ItemIndex:=0 ; Caption:=HDiv.Mess[21] ;
                  END ;
     smpDecTous : BEGIN //
                  InitDec ; CATEGORIE.Value:='' ; Caption:=HDiv.Mess[13] ;
                  END ;
   END ;
gszCaption := Caption;  // Pour les messages
PreciseMP ;
UpdateCaption(Self) ;
END ;

procedure TFMulSuiviMP.InitCriteres ;
Var i : integer ;
    StV8,St : String ;
//    Ind : tProfilTraitement ;
//    OkOk : Boolean ;
BEGIN
if ((smp<>smpEncTous) and (smp<>smpDecTous)) then
   BEGIN
   LibellesTableLibre(PzLibre,'TT_TABLE','T_TABLE','T') ;
   if VH^.CPExoRef.Code<>'' then
      BEGIN
      E_EXERCICE.Value:=VH^.CPExoRef.Code ;
      E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
      E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
      END else
      BEGIN
      E_EXERCICE.Value:=VH^.Entree.Code ;
      E_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
      E_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
      END ;
   E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   (*
   PBanqueLot.TabVisible:=False ;
   If (smp=smpDecChqEdt) or (smp=smpDecVirEdt) or (smp=smpDecVirInEdt) or (smp=smpEncPreEdt)Then
     BEGIN
     PBanqueLot.TabVisible:=TRUE ;
     END ;
   *)
   Pages.ActivePage:=PCritere ;
   END else
   BEGIN
   Pages.ActivePage:=PBanqueLot ;
   for i:=0 to Pages.PageCount-1 do if Pages.Pages[i]<>PBanqueLot then Pages.Pages[i].TabVisible:=False ;
   END ;
// ParmsMP
St:='E_ECRANOUVEAU="N"' ;
StV8:=LWhereV8 ;
if StV8<>'' then
  BEGIN
  St:='('+St+' OR E_ECRANOUVEAU="H")  ' ;
  St:=St+' AND ('+StV8+') ' ; //QEnc.SQL.Add(' AND '+StV8) ;
  END ;
XX_WHEREAN.Text:=St ;
ReinitWhereNatCpt(FALSE) ;
XX_WHEREPROFIL.Text:='' ;
(*
Ind:=QuelProfil(smp,OkOk) ; If OkOk Then BEGIN St:=WhereProfilUser(Q,Ind) ; XX_WHEREPROFIL.Text:=St ; END ;
*)
END ;

(*================================ CRITERES ==================================*)
procedure TFMulSuiviMP.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulSuiviMP.CATEGORIEChange(Sender: TObject);
var
    HCAT : THValComboBox ;
begin
    inherited;

{ Le 28/01/2003 BPY correction des bug de la fiche 11794 }
    if (Categorie.Value = 'TRI') then
    begin
        BCtrlRib.visible := false;
//        BCtrlRib.Hint := 'Contrôler les IBAN de la liste';
//    end
//    else
//    begin
//        BCtrlRib.Hint := 'Contrôler les RIB de la liste';
    end;
{ Fin BPY }

    HCAT:=THValCombobox(Sender) ;
    LaCategorieChange(HCat,XX_WHEREMP,SorteLettre) ;
    //If smp in [smpEncDiv,smpDecDiv,smpEncTraEdt,smpEncTraEdtNC] Then
    PreciseMP ;
end;

procedure TFMulSuiviMP.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
{$IFNDEF IMP}
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction)
                                  else TrouveEtLanceSaisie(Q,taConsult,'N') ;
{$ENDIF}
end;

procedure TFMulSuiviMP.ClickModifRib ;
Var RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
If ModifRibSurMul(Q,FTID.Checked,TRUE) Then
   BEGIN
   Application.ProcessMessages ; BChercheClick(Nil) ;
   RJal:=Q.FindField('E_JOURNAL').AsString ;
   RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
   RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
   RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
   RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
   RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
   Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
            VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[])
   END ;

end;

procedure TFMulSuiviMP.PrechargeOrigines ;
Var QQ : TQuery ;
    st : String ;
BEGIN
TobMvt.ClearDetail ;
St:='SELECT E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE,E_QUALIFPIECE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '+
    'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
QQ:=OpenSQL(St+RecupWhereCritere(Pages),True) ;
TobMvt.LoadDetailDB('ECRITURE','','',QQ,False,True) ;
Ferme(QQ) ;
if TobMvt.Detail.Count>0 then TobMvt.Detail[0].AddChampSup('MARQUE',True) ;
END ;

procedure TFMulSuiviMP.MarqueOrigine ;
Var TOBL : TOB ;
BEGIN
TOBL:=TobMvt.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                        [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
                         Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
                         Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
if TOBL<>Nil then TOBL.PutValue('MARQUE','X') ;
END ;

procedure TFMulSuiviMP.ToutMarquer ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TobMvt.Detail.Count-1 do
    BEGIN
    TOBL:=TobMvt.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    END ;
END ;

procedure TFMulSuiviMP.VireInutiles ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=TobMvt.Detail.Count-1 downto 0 do
    BEGIN
    TOBL:=TobMvt.Detail[i] ;
    if TOBL.GetValue('MARQUE')<>'X' then BEGIN TOBL.Free ; {TOBL:=Nil ;} END ;
    END ;
END ;

Function TFMulSuiviMP.ConstitueTobMvt : Boolean ;
Var i : integer ;
BEGIN
Result:=TRUE ;
PrechargeOrigines ;
// ============== MODE SELECTION INVERSEE
	if swapSelect then
		// Nouveau code : ajout mode selection inversée
  	begin
    Q.First ;
    While not Q.Eof do
      begin
      if  not FListe.IsCurrentSelected then MarqueOrigine ;
      Q.Next ;
      end ;
    end
  else
		// Ancien code : avant ajour mode selection inversée
  	begin
		if Not FListe.AllSelected then
      begin
      for i:=0 to FListe.NbSelected-1 do
        begin
        FListe.GotoLeBookmark(i) ;
        MarqueOrigine ;
        end ;
      end
    else
    	ToutMarquer ;
    end ;
// ============== MODE SELECTION INVERSEE
VireInutiles ;
END ;

procedure TFMulSuiviMP.ValideLot ;
Var i : Integer ;
    TobL : TOB ;
    St : String ;
BEGIN
InitMove(TobMvt.Detail.Count,'') ;
for i:=TobMvt.Detail.Count-1 downto 0 do
    BEGIN
    MoveCur(FALSE) ;
    TOBL:=TobMvt.Detail[i] ;
    St:='UPDATE ECRITURE SET E_NOMLOT="'+CodeLot+'" WHERE E_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'" '
       +' AND E_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'" AND E_DATECOMPTABLE="'+UsDateTime(TOBL.GetValue('E_DATECOMPTABLE'))+'" '
       +' AND E_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+' AND E_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))
       +' AND E_NUMECHE='+IntToStr(TOBL.GetValue('E_NUMECHE'))
       +' AND E_QUALIFPIECE="'+TOBL.GetValue('E_QUALIFPIECE')+'" ' ;
    ExecuteSQL(St) ;
    END ;
FiniMove ;
END ;

procedure TFMulSuiviMP.BOuvrirClick(Sender: TObject);
Var Abrege,Libre : String ;
begin
// Message avertissement si sélection inversée
	if swapSelect then
  	if PGIAsk('Vous êtes en mode "sélection inversée". Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes
    then  exit;
//
CodeLot:='' ; Abrege:='' ; Libre:='' ;
Abrege:=AttribFiltre(Smp) ;
//CodeLot:=AttribSoucheLotPreparation(smp) ;
If SaisieCodelot(CodeLot,smp,Abrege,'') Then
  BEGIN
  ConstitueTobMvt ;
  if Transactions(ValideLot,1)<>oeOk then
     BEGIN
     MessageAlerte(HDiv.Mess[32]) ;
     END ;
  BChercheClick(Nil) ;
  END ;
end;

procedure TFMulSuiviMP.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(Pages) ;
  If (StWRib = '') Then Exit;
  if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then begin  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do begin
      FListe.GotoLeBookmark(i) ;
      StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
    end;
    // Efface le dernier OR et rajoute ')'
    delete(StWRib,length(StWRib)-2,3);
    StWRib := StWRib +')';
  end;
  if (Categorie.Value = 'TRI') then
    CPLanceFiche_VerifRib('IBAN;WHERE='+StWRib)
  else
    CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TFMulSuiviMP.FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
Var sRIB,St : String ;
//    RJal,RExo,RQualif : String ;
//    RDate : TDateTime ;
//    RNumP,RNumL,RNumEche : Integer ;
//    OkLocate : Boolean ;
//    TOBE     : TOB ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
If FTid.Checked Then St:='E_GENERAL' Else St:='E_AUXILIAIRE' ;
if ((Field.FieldName=St) and (Q.FindField('E_RIB')<>Nil)) then
   BEGIN
   sRIB:=Q.FindField('E_RIB').AsString ; if sRib<>'' then Exit ;
   FListe.Canvas.Brush.Color:= clRed ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clRed ;
   FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
   FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
   END ;
(*
If (VH^.MPMODIFFAITE) Then
  BEGIN
  OkLocate:=TRUE ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPExoPop:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))) ;
  If Q.FindField('E_GENERAL')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPGenPop:=Q.FindField('E_GENERAL').AsString ;
  If Q.FindField('E_AUXILIAIRE')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPAuxPop:=Q.FindField('E_AUXILIAIRE').AsString ;
  If Q.FindField('E_JOURNAL')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPJalPop:=Q.FindField('E_JOURNAL').AsString ;
  If Q.FindField('E_NUMEROPIECE')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPNumPop:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  If Q.FindField('E_NUMLIGNE')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPNumLPop:=Q.FindField('E_NUMLIGNE').AsInteger ;
  If Q.FindField('E_NUMECHE')<>NIL Then OkLocate:=OkLocate And (VH^.SauveMPPop.MPNumEPop:=Q.FindField('E_NUMECHE').AsInteger ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then OkLocate:=OkLocate And (VH^.MPPop.MPDatePop:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  VH^.MPMODIFFAITE:=FALSE ;
  RJal:=Q.FindField('E_JOURNAL').AsString ;
  RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  RQualif:='N' ;
  RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
  RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
  Application.ProcessMessages ; BChercheClick(Nil) ;
  If OkLocate Then Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                            VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[]) ;
  END ;
*)
end;

procedure TFMulSuiviMP.E_NUMTRAITECHQChange(Sender: TObject);
begin
  inherited;
(*
if smp<>smpEncTraEdtNC then Exit ;
if ((E_NUMTRAITECHQ.Text<>'') or (E_NUMTRAITECHQ_.Text<>''))
   then XX_WHERETRACHQ.Text:='' else XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ=""' ;
*)
end;

procedure TFMulSuiviMP.cFactCreditClick(Sender: TObject);
begin
  inherited;
if isEncMP(smp) then
   BEGIN
   (*
   if cFactCredit.Checked then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="FC"))'
                          else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
   *)
   if cFactCredit.Checked
        then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"))'
        else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
   END else
   BEGIN
   (*
   if cFactCredit.Checked then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="FF"))'
                          else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
   *)
   if cFactCredit.Checked
      then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"))'
      else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
   END ;
end;

procedure TFMulSuiviMP.BNouvRechClick(Sender: TObject);
begin
  inherited;
CategorieChange(CATEGORIE) ;
end;

procedure TFMulSuiviMP.FListeRowEnter(Sender: TObject);
{Var RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;}
begin
  inherited;
(*
If (VH^.MPMODIFFAITE) And (FListe.Focused) Then
  BEGIN
  RJal:=Q.FindField('E_JOURNAL').AsString ;
  RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  RQualif:='N' ;
  RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
  RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
  Application.ProcessMessages ; BChercheClick(Nil) ;
  Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[]) ;
  END ;
*)
VH^.MPModifFaite:=FALSE ;
If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPExoPop:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
If Q.FindField('E_GENERAL')<>NIL Then VH^.MPPop.MPGenPop:=Q.FindField('E_GENERAL').AsString ;
If Q.FindField('E_AUXILIAIRE')<>NIL Then VH^.MPPop.MPAuxPop:=Q.FindField('E_AUXILIAIRE').AsString ;
If Q.FindField('E_JOURNAL')<>NIL Then VH^.MPPop.MPJalPop:=Q.FindField('E_JOURNAL').AsString ;
If Q.FindField('E_NUMEROPIECE')<>NIL Then VH^.MPPop.MPNumPop:=Q.FindField('E_NUMEROPIECE').AsInteger ;
If Q.FindField('E_NUMLIGNE')<>NIL Then VH^.MPPop.MPNumLPop:=Q.FindField('E_NUMLIGNE').AsInteger ;
If Q.FindField('E_NUMECHE')<>NIL Then VH^.MPPop.MPNumEPop:=Q.FindField('E_NUMECHE').AsInteger ;
If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPDatePop:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
//VH^.SauveMPPOP:=VH^.MPPOP ;
end;

procedure TFMulSuiviMP.BChercheClick(Sender: TObject);
begin
ReinitWhereNatCpt(FALSE) ;
If Trim(E_GENERAL.Text)<>'' Then ReinitWhereNatCpt(TRUE) ;
If Trim(E_AUXILIAIRE.Text)<>'' Then ReinitWhereNatCpt(TRUE) ;
inherited;
AfficheSoldeSelection ;
FListeRowEnter(nil); // 10930
end;

procedure TFMulSuiviMP.FTIDClick(Sender: TObject);
begin
  inherited;
E_GENERAL.Text:='' ;
If FTID.Checked Then
  BEGIN
  E_AUXILIAIRE.Enabled:=FALSE ; TE_AUXILIAIRE.Enabled:=FALSE ;
  E_AUXILIAIRE.Text:='' ;
  If IsEnc(smp) Then E_GENERAL.DataType:='TZGTID' Else E_GENERAL.DataType:='TZGTIC' ;
  END Else
  BEGIN
  E_AUXILIAIRE.Enabled:=TRUE ; TE_AUXILIAIRE.Enabled:=TRUE ;
  If IsEnc(smp) Then E_GENERAL.DataType:='TZGENCAIS' Else E_GENERAL.DataType:='TZGDECAIS';
  END ;
end;

procedure TFMulSuiviMP.TraiteEditeClick(Sender: TObject);
begin
  inherited;
E_NUMTRAITECHQ_.Enabled:=TraiteEdite.State=cbChecked ;
FE_NUMTRAITECHQ_.Enabled:=TraiteEdite.State=cbChecked ;
E_NUMTRAITECHQ.Enabled:=TraiteEdite.State=cbChecked ;
FE_NUMTRAITECHQ.Enabled:=TraiteEdite.State=cbChecked ;
XX_WHERETRACHQ.Text:='' ;
If TraiteEdite.State=cbGrayed Then
  BEGIN
  E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
  XX_WHERETRACHQ.Text:='' ;
  END ;
If TraiteEdite.State=cbChecked Then
  BEGIN
  E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
  XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ<>""' ;
  END ;
If TraiteEdite.State=cbUnChecked Then
  BEGIN
  E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
  XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ=""' ;
  END ;
end;

procedure TFMulSuiviMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//Var S : Boolean ;
begin
  inherited;
//S:=(Shift=[ssCtrl]) ;
Case Key of
  VK_F5 : BEGIN Key:=0 ; ClickModifRib ; END
  END ;
end;

{procedure TFMulSuiviMP.FListeEnter(Sender: TObject);
Var i : Integer ;
begin
  inherited;
i:=0 ;
end;}

procedure TFMulSuiviMP.FormClose(Sender: TObject; var Action: TCloseAction);
{Var Ind : tProfilTraitement ;
    OkOk : Boolean ;}
begin
  inherited;
TobMvt.Free ;
//Ind:=QuelProfil(smp,OkOk) ;
(*If Not MonProfilOk(Ind) Then*)
(* GP le 18/12/2001 Bloqueur(smpRes,False) ; *)
end;

Function TFMulSuiviMP.GereXNomLot : Boolean ;
Var St,XAnd,St1,St2 : String ;
BEGIN
Result:=FALSE ;
XX_WHERELOT.Text:=' E_NOMLOT="" ' ; St1:='' ; St2:='' ;
If XNOMLOT1.Text<>'' Then St1:=' (E_NOMLOT>="'+XNOMLOT1.Text+'") ' ;
If XNOMLOT2.Text<>'' Then St2:=' (E_NOMLOT<="'+XNOMLOT2.Text+'") ' ;
If (St1<>'') And (St2<>'') Then XAnd:=' AND ' Else XAnd:='' ;
If (St1<>'') Or (St2<>'') Then
  BEGIN
  Result:=TRUE ;
  If Not CRechVide.Checked Then St:=' (E_NOMLOT<>"" AND ('+St1+XAnd+St2+')) '
                       Else St:=' (E_NOMLOT="" OR ('+St1+XAnd+St2+')) ' ;
  XX_WhereLot.Text:=St ;
  END ;
END ;

procedure TFMulSuiviMP.XNOMLOT1Change(Sender: TObject);
begin
  inherited;
If GereXNomLot Then
  BEGIN
  CRechVide.Enabled:=TRUE ;
  END Else
  BEGIN
  CRechVide.Enabled:=FALSE ;
  CRechVide.Checked:=FALSE ;
  END ;
end;

procedure TFMulSuiviMP.CRechVideClick(Sender: TObject);
begin
  inherited;
(*If CRechVide.Checked Then*) GereXNomLot ;
end;

procedure TFMulSuiviMP.BRAZLotClick(Sender: TObject);
begin
  inherited;
If HDiv.Execute(33,gszCaption,'')<>mrYes Then Exit ;
CodeLot:='' ;
ConstitueTobMvt ;
if Transactions(ValideLot,1)<>oeOk then
   BEGIN
   MessageAlerte(HDiv.Mess[32]) ;
   END ;
BChercheClick(Nil) ;
end;

procedure TFMulSuiviMP.NewFiltreMPClick(Sender: TObject);
begin
  inherited;
VideFiltre(FFiltres,Pages) ; 
end;

procedure TFMulSuiviMP.bSelectAllClick(Sender: TObject);
begin
  inherited;
  // Si on sélectionne tout, on annule le mode "sélection inversée"
  if FListe.AllSelected then
  	begin
    SwapSelect := False ;
    BSwapSelect.Down := False ;
		BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
    end;
	AfficheSoldeSelection ;
end;

procedure TFMulSuiviMP.BKillLotClick(Sender: TObject);
begin
  inherited;
DetruitLeLot(smp) ;
end;

{ BPY 29/01/2003 Correction de la fiche 11794
procedure TFMulSuiviMP.DessineCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  Etab,Guichet,Numero,Cle,Dom,Iban,Rib : String ;
begin
  inherited;
  If Assigned(Column.Field) and (Column.FieldName ='E_RIB') and (not Column.Field.IsNull) then begin
    if (Categorie.Value = 'TRI') then begin
      Rib := Column.Field.Value;
      DecodeRIBIban(Etab,Guichet,Numero,Cle,Dom,Iban,Rib);
      FListe.Canvas.FillRect(Rect);
      FListe.Canvas.TextRect (Rect,Rect.Left+2,Rect.Top,Iban);
      Column.Title.Caption := 'IBAN';
    end;
  end;
end;
}

procedure TFMulSuiviMP.AfficheSoldeSelection;
Var Solde : Double;
begin
	// Calcul du solde
	Solde := CalculSoldeSelection;
  // Init du titre
//  InitTitreContexte;
  // Ajout affichage du cumul
  if Solde < 0
  	then Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(-1*Solde,0,V_PGI.OkDecV,'',True) + ' C)'
  	else Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(Solde,0,V_PGI.OkDecV,'',True) + ' D)';
	// Affichage de mode "selection inversé" si besoin
  if SwapSelect then
  	Caption := Caption + ' [SELECTION INVERSEE] ';
	// Raffraichissement Titre
	UpdateCaption(Self) ;
end;

function TFMulSuiviMP.CalculSoldeSelection: Double;
Var i 										: Integer ;
		total,debit, credit  	: Double ;
begin
	Result := 0 ;
  if (PCumul.FindChildControl('__QRYPCumul_E_DEBIT') = nil)
  	 or (PCumul.FindChildControl('__QRYPCumul_E_CREDIT') = nil) then Exit ;
  if FListe.AllSelected then	// Si tout sélectionné, on parcours le Query
    begin
	  debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text ) ;
  	credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text ) ;
	  Result := debit - credit ;
    end
  else	// Sinon, on parcours le Bookmark
    begin
    for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
      Result := Result + Q.FindField('E_DEBIT').asFloat - Q.FindField('E_CREDIT').asFloat ;
      END ;
    if SwapSelect then  // en mode inversé, on retranche le cumul des lignes sélectionnée du solde total
      begin
		  debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text ) ;
  		credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text ) ;
	  	Total	 := debit - credit ;
      if ((Result < 0) and (Total<0)) or ((Result >= 0) and (Total >= 0))
         then Result := Total - Result
         else Result := Total + Result ;
      end ;
    end ;
end;

{procedure TFMulSuiviMP.InitTitreContexte;
begin
	Caption := '?? -';
	Case smp of
  	smpEncDiv    		: Caption:=HDiv.Mess[20] ;
   	smpEncPreEdt 		: Caption:=HDiv.Mess[19] ;
   	smpEncPreBqe 		: Caption:=HDiv.Mess[29] ;
   	smpEncChqPor 		: Caption:=HDiv.Mess[28] ;
   	smpEncCBPor 		: Caption:=HDiv.Mess[16] ;
   	smpEncTraPor 		: Caption:=HDiv.Mess[16] ;
   	smpEncTraEdt 		: Caption:=HDiv.Mess[5] ;
 		smpEncTraEdtNC 	: Caption:=HDiv.Mess[26] ;
		smpEncTraEnc 		: Caption:=HDiv.Mess[8] ;
   	smpEncTraEsc 		: Caption:=HDiv.Mess[9] ;
		smpEncChqBqe 		: Caption:=HDiv.Mess[31] ;
		smpEncCBBqe 		: Caption:=HDiv.Mess[6] ;
		smpEncTraBqe 		: Caption:=HDiv.Mess[6] ;
    smpEncTous 			: Caption:=HDiv.Mess[12] ;
		smpDecChqEdt 		: Caption:=HDiv.Mess[2] ;
		smpDecChqEdtNC 	: Caption:=HDiv.Mess[2] ;
		smpDecVirEdt 		: Caption:=HDiv.Mess[18] ;
		smpDecVirEdtNC 	: Caption:=HDiv.Mess[18] ;
		smpDecVirBqe 		: Caption:=HDiv.Mess[3] ;
	  smpDecVirInEdt 	: Caption:=HDiv.Mess[35] ;
	  smpDecVirInEdtNC: Caption:=HDiv.Mess[35] ;
    smpDecVirInBqe	: Caption:=HDiv.Mess[36] ;
		smpDecBorEdt 		: Caption:=HDiv.Mess[27] ;
 		smpDecBorEdtNC 	: Caption:=HDiv.Mess[4] ;
   	smpDecborDec 		: Caption:=HDiv.Mess[10] ;
  	smpDecBorEsc 		: Caption:=HDiv.Mess[11] ;
		smpDecTraBqe 		: Caption:=HDiv.Mess[7] ;
    smpDecTraPor 		: Caption:=HDiv.Mess[17] ;
    smpDecDiv    		: Caption:=HDiv.Mess[21] ;
    smpDecTous 			: Caption:=HDiv.Mess[13] ;
  END ;
end;}

procedure TFMulSuiviMP.FListeFlipSelection(Sender: TObject);
begin
  inherited;
	AfficheSoldeSelection ;
end;

procedure TFMulSuiviMP.BSwapSelectClick(Sender: TObject);
begin
  // Si tout sélectionné, on déselectionne tout
  if FListe.AllSelected then bSelectAllClick(nil) ;
  // Inverse la sélection
  SwapSelect := not SwapSelect ;
  // Modification du Hint
  if SwapSelect
  	then BSwapSelect.Hint := 'Désactiver le mode "sélection inversée"'
  	else BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
  // Recalcul du solde
 	AfficheSoldeSelection ;
end;

end.

