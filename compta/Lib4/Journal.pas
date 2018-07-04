unit Journal;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, StdCtrls, Tabs,
     Buttons, ExtCtrls, Grids, Mask,  DB, {$IFNDEF DBXPRESS}dbtables,
  HSysMenu, Menus, Hqry, hmsgbox, DBCtrls, ComCtrls, HRichEdt, HRichOLE,
  Hcompte, HRegCpte, Hctrls, ADODB, udbxDataset, TntComCtrls, TntStdCtrls,
  TntButtons{$ELSE}uDbxDataSet{$ENDIF}, DBCtrls, Hctrls, HDB,
     DBGrids, Hqry, TabNotBk, Dialogs, Spin, ENT1,  Hcompte,
     SysUtils, HmsgBox, ComCtrls, Hent1, Messages, Menus, Filtre,
     HRichEdt, HSysMenu, HRichOLE, MajTable, HRegCpte, ParamSoc   ,UentCommun;

Procedure FicheJournal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

type
  TFJournal = class(TForm)
    SJal  : TDataSource;
    DBNav : TDBNavigator;
    HPB   : TPanel;
    Q1    : THQuery;
    Q2    : THQuery;
    BImprimer: THBitBtn;
    BAnnuler: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BFirst: THBitBtn;
    BPrior: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    Binsert: THBitBtn;
    BAide: THBitBtn;
    MsgBox        : THMsgBox;
    Pages         : TPageControl;
    PCar          : TTabSheet;
    PInf          : TTabSheet;
    PComplement   : TTabSheet;
    PReg          : TTabSheet;
    HGBDERNMOUV         : TGroupBox;
    TJ_DEBITDERNMVT     : THLabel;
    TJ_CREDITDERNMVT    : THLabel;
    TJ_NUMDERNMVT       : THLabel;
    J_DEBITDERNMVT      : TDBEdit;
    J_CREDITDERNMVT     : TDBEdit;
    J_NUMDERNMVT        : TDBEdit;
    TJ_COMPTEINTERDIT   : THLabel;
    J_COMPTEINTERDIT    : TDBEdit;
    TJ_COMPTEAUTOMAT    : THLabel;
    J_COMPTEAUTOMAT     : TDBEdit;
    J_COMPTEURNORMAL    : THDBValComboBox;
    TJ_COMPTEURNORMAL   : THLabel;
    TJ_COMPTEURSIMUL    : THLabel;
    J_COMPTEURSIMUL     : THDBValComboBox;
    J_CONTREPARTIE      : THDBCpteEdit;
    TJ_CONTREPARTIE     : THLabel;
    TJ_TYPECONTREPARTIE : THLabel;
    J_TYPECONTREPARTIE  : THDBValComboBox;
    TJ_BLOCNOTE         : TGroupBox;
    J_BLOCNOTE          : THDBRichEditOLE;
    HGBDates            : TGroupBox;
    TJ_DATEOUVERTURE    : THLabel;
    TJ_DATEFERMETURE    : THLabel;
    TJ_DATEMODIFICATION : THLabel;
    TJ_DATECREATION     : THLabel;
    J_FERME             : TDBCheckBox;
    J_DATEDERNMVT       : TDBEdit;
    J_DATECREATION      : TDBEdit;
    J_DATEOUVERTURE     : TDBEdit;
    J_DATEMODIF: TDBEdit;
    J_DATEFERMETURE     : TDBEdit;
    FAutoSave           : TCheckBox;
    TJ_JOURNAL          : THLabel;
    J_JOURNAL           : TDBEdit;
    TJ_NATUREJAL        : THLabel;
    J_NATUREJAL         : THDBValComboBox;
    TJ_LIBELLE          : THLabel;
    J_LIBELLE           : TDBEdit;
    TJ_ABREGE           : THLabel;
    J_ABREGE            : TDBEdit;
    J_MULTIDEVISE       : TDBCheckBox;
    J_CENTRALISABLE     : TDBCheckBox;
    TJ_AXE              : THLabel;
    J_AXE               : THDBValComboBox;
    GBCarac             : TGroupBox;
    HLabel1             : THLabel;
    HLabel3             : THLabel;
    HLabel2             : THLabel;
    HLabel4             : THLabel;
    GbReg               : TGroupBox;
    TJ_CPTEREGULDEBIT1  : TLabel;
    J_CPTEREGULDEBIT1   : THDBCpteEdit;
    J_CPTEREGULCREDIT1  : THDBCpteEdit;
    TJ_CPTEREGULDEBIT2  : TLabel;
    J_CPTEREGULDEBIT2   : THDBCpteEdit;
    J_CPTEREGULCREDIT2  : THDBCpteEdit;
    TJ_CPTEREGULDEBIT3  : TLabel;
    J_CPTEREGULDEBIT3   : THDBCpteEdit;
    J_CPTEREGULCREDIT3  : THDBCpteEdit;
    HLabel5             : THLabel;
    J_TOTDEBP           : TDBEdit;
    J_TOTCREP           : TDBEdit;
    JSOLCREP            : THNumEdit;
    HLabel6             : THLabel;
    J_TOTDEBE           : TDBEdit;
    J_TOTCREE           : TDBEdit;
    JSOLCREE            : THNumEdit;
    J_TOTDEBS           : TDBEdit;
    J_TOTCRES           : TDBEdit;
    JSOLCRES            : THNumEdit;
    HLabel7             : THLabel;
    PopZ                : TPopupMenu;
    BMenuZoom: THBitBtn;
    BJALDIV: THBitBtn;
    BJALCPT: THBitBtn;
    BJALPER: THBitBtn;
    BJALCENT: THBitBtn;
    BZecrimvt: THBitBtn;
    BCumul: THBitBtn;
    QJal: TQuery;
    QJalJ_JOURNAL          : TStringField;
    QJalJ_LIBELLE          : TStringField;
    QJalJ_ABREGE           : TStringField;
    QJalJ_NATUREJAL        : TStringField;
    QJalJ_DATECREATION     : TDateTimeField;
    QJalJ_DATEOUVERTURE    : TDateTimeField;
    QJalJ_DATEFERMETURE    : TDateTimeField;
    QJalJ_FERME            : TStringField;
    QJalJ_DATEDERNMVT      : TDateTimeField;
    QJalJ_DEBITDERNMVT     : TFloatField;
    QJalJ_CREDITDERNMVT    : TFloatField;
    QJalJ_NUMDERNMVT       : TIntegerField;
    QJalJ_BLOCNOTE         : TMemoField;
    QJalJ_COMPTEURNORMAL   : TStringField;
    QJalJ_COMPTEURSIMUL    : TStringField;
    QJalJ_MULTIDEVISE      : TStringField;
    QJalJ_TYPECONTREPARTIE : TStringField;
    QJalJ_CONTREPARTIE     : TStringField;
    QJalJ_COMPTEINTERDIT   : TStringField;
    QJalJ_CENTRALISABLE    : TStringField;
    QJalJ_UTILISATEUR      : TStringField;
    QJalJ_TOTALDEBIT       : TFloatField;
    QJalJ_TOTALCREDIT      : TFloatField;
    QJalJ_CPTEREGULDEBIT1  : TStringField;
    QJalJ_CPTEREGULDEBIT2  : TStringField;
    QJalJ_CPTEREGULDEBIT3  : TStringField;
    QJalJ_CPTEREGULCREDIT1 : TStringField;
    QJalJ_CPTEREGULCREDIT2 : TStringField;
    QJalJ_CPTEREGULCREDIT3 : TStringField;
    QJalJ_SOCIETE          : TStringField;
    QJalJ_COMPTEAUTOMAT    : TStringField;
    QJalJ_TOTDEBP          : TFloatField;
    QJalJ_TOTCREP          : TFloatField;
    QJalJ_TOTDEBE          : TFloatField;
    QJalJ_TOTCREE          : TFloatField;
    QJalJ_TOTDEBS          : TFloatField;
    QJalJ_TOTCRES          : TFloatField;
    QJalJ_AXE              : TStringField;
    QJalJ_VALIDEEN         : TStringField;
    QJalJ_VALIDEEN1        : TStringField;
    HMTrad: THSystemMenu;
    Bevel2: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    QJalJ_DATEMODIF: TDateTimeField;
    QJalJ_EFFET: TStringField;
    J_EFFET: TDBCheckBox;
    QJalJ_MODESAISIE: TStringField;
    TJ_MODESAISIE: THLabel;
    J_MODESAISIE: THDBValComboBox;
    PTreso: TTabSheet;
    HLabel9: THLabel;
    J_TRESODATE: THDBValComboBox;
    QJalJ_CHOIXDATE: TStringField;
    PSAISIE: TTabSheet;
    J_INCREF: TDBCheckBox;
    J_NATCOMPL: TDBCheckBox;
    J_EQAUTO: TDBCheckBox;
    J_INCNUM: TDBCheckBox;
    J_NATDEFAUT: THDBValComboBox;
    HLabel8: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPriorClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure SJalDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFermeClick(Sender: TObject);
    procedure J_NATUREJALChange(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BCumulClick(Sender: TObject);
    procedure J_CPTEREGULDEBIT1Exit(Sender: TObject);
    procedure QJalAfterDelete(DataSet: TDataSet);
    procedure QJalAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BZecrimvtClick(Sender: TObject);
    procedure BMenuZoomClick(Sender: TObject);
    procedure BJALDIVClick(Sender: TObject);
    procedure BJALCPTClick(Sender: TObject);
    procedure BJALPERClick(Sender: TObject);
    procedure BJALCENTClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
    procedure J_EFFETClick(Sender: TObject);
    procedure J_MODESAISIEChange(Sender: TObject);
    procedure J_JOURNALExit(Sender: TObject);
    procedure PagesChanging(Sender: TObject; var AllowChange: Boolean);
    procedure J_JOURNALChange(Sender: TObject);
  private
    Q : TQuery ;
    QSouche : TQuery ;
    Lequel,LeCompte : String ;
    Mode   : TActionFiche ;
    LaPage : Integer ;
    AvecMvt: Boolean ;
    FAvertir : Boolean ;
    QJalJ_INCNUM: TStringField ;
    QJalJ_INCREF: TStringField ;
    QJalJ_NATCOMPL: TStringField ;
    QJalJ_EQAUTO: TStringField ;
    QJalJ_NATDEFAUT: TStringField ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  CodExist : Boolean ;
    Function  Supprime        : Boolean ;
    Function  EstJalAnal : Boolean ;
    Procedure AutoriseOuNonLesChamps ;
    Procedure IsMouvemente ;
    (*P.FUGIER 05/00 : DEBUT*)
    procedure InitFolio ;
    (*P.FUGIER 05/00 : FIN*)
    Procedure ClearChampsSurEnregOk ;
    Function  DansUnCpte(Cpte1,Cpte2 : TObject) : Boolean ;
    Function  CpteDeRegulValide(Sender : TObject) : Boolean ;
    Function  CrtlCpteDeRegul : Boolean ;
    Function  ChaineCpteValide(Sender : TObject) : Boolean ;
    Function  ExisteLeCompte(LeWhere : String) : Boolean ;
    Function  CaractereValide(Sender : TObject) : Boolean ;
    Function  FaitLeWhere(St : String) : String ;
    Procedure MsgCpteRegul(Sender : TOBject ; i : Byte) ;
    function  ContrePartieValide : Boolean;
    function  ContrePartieDevise : Boolean;
    Procedure DevalideLexoSurCreatJal ;
    Procedure FormateLesMontants ;
    Procedure ActivationControl;
  public
  end;

    Function  EstDansAnalytiq(St : String) : Boolean ;
    Function  EstDansEcriture(St : String) : Boolean ;
    Function  EstDansSociete(St : String)  : Boolean ;
    Function  EstDansSouche(St : String)   : Boolean ;
    Function  EstDansPiece(St : String)   : Boolean ;           // Modif Fiche 10856 SBO
    Function  EstDansParamPiece(St : String)   : Boolean ;      // Modif Fiche 10856 SBO
    Function  EstDansParamPieceCompl(St : String)   : Boolean ; // Modif Fiche 10856 SBO


implementation
{$R *.DFM}

uses
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
     QRPlanJo , CumMens, Zoomana, Zecrimvt,
{$ENDIF}
{$ENDIF}
     SaisUtil, SaisComm,
{$IFNDEF PGIIMMO}
     UTILEDT,
{$ENDIF}
     CPTESAV,
     uLibWindows,UFonctionsCBP;

Procedure FicheJournal(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var Fjournal: TFjournal ;
begin
Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not ExJaiLeDroitConcept(TConcept(ccJalCreat),True) then Exit ;
              taModif,taModifEnSerie : if Not ExJaiLeDroitConcept(TConcept(ccJalModif),True) then Exit ;
   END ;
FJournal:=TFjournal.Create(Application) ;
try
  FJournal.Q:=Q ;
  FJournal.Lequel:=Compte ;
  FJournal.Mode:=Comment ;
  FJournal.LaPage:=QuellePage ;
  FJournal.ShowModal ;
 finally
  FJournal.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/10/2002
Modifié le ... :   /  /    
Description .. : - 21/10/2002 - le panel de saisie des info bordereau n'est 
Suite ........ : visible que pour un bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFJournal.FormShow(Sender: TObject);
begin
MakeZoomOLE(Handle) ;
RecupWhereSQL(Q,QJal) ;
if ((Q=NIL) and (Lequel<>'')) then QJal.SQL.Add('Where J_JOURNAL="'+Lequel+'"') ;
if (Q=NIL) then QJal.SQL.Add('Order by J_JOURNAL') ;//lhe le 19/02/97 pour trier [tacreat..tacreatenserie] avec navigation
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QJal) ;
ChangeSizeMemo(QJalJ_BLOCNOTE) ;
QJal.Open ; FormateLesMontants ; Pages.ActivePage:=Pages.Pages[LaPage] ;
if (not V_PGI.MonoFiche) and (Lequel<>'') And ((Mode in [taCreat..taCreatOne])=False) then
  BEGIN
  if Not QJal.Locate('J_JOURNAL',Lequel,[]) then
    BEGIN MessageAlerte(MsgBox.Mess[18]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
  END ;
Case Mode of
     taConsult : BEGIN
                 FicheReadOnly(Self) ;
                 PTreso.TabVisible := ( J_NATUREJAL.Value = 'BQE' ) and (ctxPCL in V_PGI.PGIContexte);
                 PSaisie.TabVisible:= (( J_MODESAISIE.Value = 'BOR' ) or ( J_MODESAISIE.Value = 'LIB' ));
                 END ;
     taCreat..taCreatOne : BEGIN Bouge(nbInsert) ; J_JOURNAL.Text:=Lequel ; BAnnuler.Enabled:=False ; END ;
   End ;
BFirst.Visible := not V_PGI.MonoFiche ;
BPrior.Visible := not V_PGI.MonoFiche ;
BNext.Visible := not V_PGI.MonoFiche ;
BLast.Visible := not V_PGI.MonoFiche ;
if SaisieFolioLancee then BEGIN J_COMPTEINTERDIT.Enabled:=False ; J_COMPTEAUTOMAT.Enabled:=False ; END ;
(*
if not ((ctxPCL in V_PGI.PGIContexte) And OkSynchro) then
  begin J_MODESAISIE.Plus:='AND CO_CODE<>"LIB"' ; J_MODESAISIE.Exhaustif:=exPlus ; end ;
{$IFDEF CCS3}
J_AXE.Visible:=False ; TJ_AXE.Visible:=False ;
// J_TYPECONTREPARTIE.Visible:=False ; TJ_TYPECONTREPARTIE.Visible:=False ;
{$ENDIF}
*)
end;

Procedure TFJournal.IsMouvemente ;
Var NatJal : String ;
BEGIN
if ((Mode=taConsult) or (QJal.State=dsInsert)) then BEGIN AvecMvt:=False ; Exit ; END ;
Natjal:=J_NATUREJAL.Value ;
if(NatJal='ODA') Or (NatJal='ANA')then
  AvecMvt:=ExisteSql('Select J_JOURNAL From JOURNAL Where J_JOURNAL="'+QJalJ_JOURNAL.AsString+'" '+
                     'And EXISTS(SELECT Y_JOURNAL FROM ANALYTIQ WHERE Y_JOURNAL="'+QJalJ_JOURNAL.AsString+'" )')
else
  AvecMvt:=ExisteSql('Select J_JOURNAL From JOURNAL Where J_JOURNAL="'+QJalJ_JOURNAL.AsString+'" '+
                     'And EXISTS(SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL="'+QJalJ_JOURNAL.AsString+'" )') ;
END ;

(*P.FUGIER 05/00 : DEBUT*)
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/10/2002
Modifié le ... :   /  /    
Description .. : - 21/10/2002 - affiche le panel des info complemetaire pour 
Suite ........ : un bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFJournal.InitFolio ;
begin
if J_MODESAISIE.Value='-'
  then begin
  J_TYPECONTREPARTIE.Enabled :=((J_NatureJal.Value='BQE') Or (J_NatureJal.Value='CAI') Or (QJalJ_EFFET.AsString='X')) ;
  TJ_TYPECONTREPARTIE.Enabled:=((J_NatureJal.Value='BQE') Or (J_NatureJal.Value='CAI') Or (QJalJ_EFFET.AsString='X')) ;
  end
  else begin
  J_TYPECONTREPARTIE.ItemIndex:=J_TYPECONTREPARTIE.Values.IndexOf('MAN') ;
  J_TYPECONTREPARTIE.Enabled:=FALSE ; TJ_TYPECONTREPARTIE.Enabled:=FALSE ;
  end ;
PSaisie.TabVisible:= (( J_MODESAISIE.Value = 'BOR' ) or ( J_MODESAISIE.Value = 'LIB' ));
end ;
(*P.FUGIER 05/00 : FIN*)

Procedure TFJournal.ChargeEnreg ;
BEGIN
if LeCompte=QJALJ_JOURNAL.AsString then
   BEGIN
   if Mode=taConsult then FicheReadOnly(Self) ;
   Exit ;
   END ;
LeCompte:=QJALJ_JOURNAL.AsString ;

InitCaption(Self,J_Journal.text,J_LIBELLE.text) ;
AfficheLeSolde(JSOLCREP,QJalJ_TOTDEBP.AsFloat,QJalJ_TOTCREP.AsFloat) ;
AfficheLeSolde(JSOLCREE,QJalJ_TOTDEBE.AsFloat,QJalJ_TOTCREE.AsFloat) ;
AfficheLeSolde(JSOLCRES,QJalJ_TOTDEBS.AsFloat,QJalJ_TOTCRES.AsFloat) ;
PReg.TabVisible:=(J_NATUREJAL.Value='REG') ;
if Mode=taConsult then
  BEGIN
  FicheReadOnly(Self) ;
  if(J_NATUREJAL.Value='ODA')Or(J_NATUREJAL.Value='ANA')then J_COMPTEURNORMAL.DataType:='ttSoucheComptaODA'
                                                        else J_COMPTEURNORMAL.DataType:='ttSoucheCompta' ;
  Exit ;
  END ;
IsMouvemente ;
J_Journal.Enabled:=False ; AutoriseOuNonLesChamps ;
END ;

Function TFJournal.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if QJal.Modified then
   BEGIN
   if (Mode in [taCreat..taCreatOne])And(J_JOURNAL.Text='')then Rep:=mrNo else
     if FAutoSave.Checked then Rep:=mrYes else Rep:=Msgbox.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes    : if not Bouge(nbPost)   then Exit ;
  mrNo     : if not Bouge(nbCancel) then Exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Procedure TFJournal.ClearChampsSurEnregOk ;
Var NatJal : String ;
    i : Byte ;
BEGIN
if QJal.State=dsBrowse then Exit ;
Natjal:=J_NATUREJAL.Value ;
if NatJal<>'REG' then
   BEGIN
   for i:=1 to 3 do
       BEGIN
       QJal.FindField('J_CPTEREGULDEBIT'+IntToStr(i)+'').AsString:='' ;
       QJal.FindField('J_CPTEREGULCREDIT'+IntToStr(i)+'').AsString:='' ;
       END ;
   END ;
if(NatJal<>'ODA') And (NatJal<>'ANA') then QJalJ_AXE.AsString:='' ;
if(NatJal='ODA') Or (NatJal='ANA')then
  BEGIN
  QJalJ_MULTIDEVISE.AsString:='-' ; QJalJ_COMPTEINTERDIT.AsString:='' ;
  QJalJ_COMPTEAUTOMAT.AsString:='' ;QJalJ_COMPTEURSIMUL.AsString:='' ;
  QJalJ_CONTREPARTIE.AsString:='' ; QJalJ_TYPECONTREPARTIE.AsString:='' ;
  Exit ;
  END ;
if(NatJal='ACH')Or(NatJal='ECC')Or(NatJal='EXT') Or((NatJal='OD') And (Not J_EFFET.Checked))Or(NatJal='REG')Or(NatJal='VTE') then
  BEGIN
  QJalJ_CONTREPARTIE.AsString:='' ; QJalJ_TYPECONTREPARTIE.AsString:='' ;
  Exit ;
  END ;
if(NatJal='ANO') then
  BEGIN
  QJalJ_COMPTEINTERDIT.AsString:='' ;
  QJalJ_COMPTEAUTOMAT.AsString:='' ;QJalJ_COMPTEURSIMUL.AsString:='' ;
  QJalJ_CONTREPARTIE.AsString:='' ; QJalJ_TYPECONTREPARTIE.AsString:='' ;
  QJalJ_CENTRALISABLE.AsString:='-';
  Exit ;
  END ;
if(NatJal='BQE')Or(NatJal='CAI') then
  BEGIN
  QJalJ_COMPTEAUTOMAT.AsString:='' ;
  Exit ;
  END ;
if(NatJal='CLO')then
  BEGIN
  QJalJ_CENTRALISABLE.AsString:='-'; QJalJ_COMPTEINTERDIT.AsString:='' ;
  QJalJ_COMPTEAUTOMAT.AsString:='' ; QJalJ_COMPTEURSIMUL.AsString:='' ;
  QJalJ_CONTREPARTIE.AsString:=''  ; QJalJ_TYPECONTREPARTIE.AsString:='' ;
  QJalJ_MULTIDEVISE.AsString:='-' ; Exit ;
  END ;
END ;

Function TFJournal.EnregOK : boolean ;
Var NatJal : String ;
BEGIN
result:=FALSE  ;
if QJal.State in [dsInsert,dsEdit]=False then Exit ;
NatJal:=J_NatureJal.Value ;
(*
if not ((ctxPCL in V_PGI.PGIContexte)) then
  BEGIN
  If ((QJal.State in [dsInsert]) Or (J_MODESAISIE.Enabled And (QJal.State in [dsEdit])))And (QJALJ_MODESAISIE.AsString='LIB') Then
    BEGIN
    Pages.ActivePage:=PCar ; J_ModeSaisie.SetFocus ; MsgBox.execute(27,'','') ; Exit ;
    END ;
  END ;
*)
if QJal.State in [dsInsert,dsEdit] then
   BEGIN
   if QJalJ_Journal.AsString='' then
      BEGIN
      Pages.ActivePage:=PCar ; J_Journal.SetFocus ; MsgBox.execute(2,'','') ; Exit ;
      END ;
   if QJalJ_LIBELLE.AsString='' then
      BEGIN
      Pages.ActivePage:=PCar ; J_LIBELLE.SetFocus ; MsgBox.execute(3,'','') ; Exit ;
      END ;
   if Not CaractereValide(J_COMPTEINTERDIT)   then Exit ;
   if Not CaractereValide(J_COMPTEAUTOMAT)    then Exit ;
   if Not CaractereValide(J_CONTREPARTIE)     then Exit ;

   if Not ChaineCpteValide(J_COMPTEINTERDIT)  then Exit ;
   if Not ChaineCpteValide(J_COMPTEAUTOMAT)   then Exit ;
   if Not ChaineCpteValide(J_CONTREPARTIE)    then Exit ;

   if NatJal='REG' then
      BEGIN
      if Not CaractereValide(J_CPTEREGULDEBIT1)  then Exit ;
      if Not CaractereValide(J_CPTEREGULCREDIT1) then Exit ;
      if Not CaractereValide(J_CPTEREGULDEBIT2)  then Exit ;
      if Not CaractereValide(J_CPTEREGULCREDIT2) then Exit ;
      if Not CaractereValide(J_CPTEREGULDEBIT3)  then Exit ;
      if Not CaractereValide(J_CPTEREGULCREDIT3) then Exit ;
      if Not CrtlCpteDeRegul then Exit ;
      END ;

   if not J_TYPECONTREPARTIE.Enabled then QJalJ_TYPECONTREPARTIE.AsString:='MAN' ;
   if((NatJal='BQE')Or(NatJal='CAI'))And((Not ContrePartieValide)Or(Not ContrePartieDevise))then Exit ;

   if((NatJal='ACH')Or(NatJal='ECC')Or(NatJal='EXT')Or
      (NatJal='REG')Or(NatJal='VTE')Or(NatJal='OD')) then
      BEGIN
      if DansUnCpte(J_COMPTEAUTOMAT,J_COMPTEINTERDIT) then
         BEGIN
         Pages.ActivePage:=PComplement ; J_COMPTEAUTOMAT.SetFocus ;
         MsgBox.Execute(8,'','') ; Exit ;
         END ;
      END ;
   if(NatJal='ECC') then QJalJ_MULTIDEVISE.AsString:='X' ;
   if(NatJal='ODA') Or (NatJal='ANA')then
      if Not EstJalAnal then Exit ;
   If QJalJ_COMPTEURNORMAL.AsString='' Then
      BEGIN
      { AJOUT ME LE 18/04/2001  MODIF ME 05/06/01}
      if ctxPCL in V_PGI.PGIContexte then
      begin
        if (NatJal='ODA') Or (NatJal='ANA') then
        begin
             if ExisteSQL ('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="ANA"') then
                QJalJ_COMPTEURNORMAL.AsString := 'ANA'
             else
             begin
                  Pages.ActivePage:=PComplement ; J_COMPTEURNORMAL.SetFocus ;
                  MsgBox.Execute(26,'','') ; Exit;
             end;
        end
        else
        begin
            if(NatJal='ANO') Or (NatJal='CLO')then
            begin
                 if ExisteSQL ('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="'+NatJal+'"') then
                 QJalJ_COMPTEURNORMAL.AsString := NatJal
                 else
                 begin
                      Pages.ActivePage:=PComplement ; J_COMPTEURNORMAL.SetFocus ;
                      MsgBox.Execute(26,'','') ; Exit;
                 end;
            end
            else
            begin
                if ExisteSQL ('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="CPT"') then
                  QJalJ_COMPTEURNORMAL.AsString := 'CPT'
                else
                begin
                  Pages.ActivePage:=PComplement ; J_COMPTEURNORMAL.SetFocus ;
                  MsgBox.Execute(26,'','') ; Exit ;
                end;
            end;
        end;
      end else   // si non  PCL
      begin
if VH^.CPIFDEFCEGID then
Begin
      QSouche:=OpenSQL('SELECT * FROM SOUCHE WHERE SH_SOUCHE="'+QJalJ_JOURNAL.AsString+'" AND SH_TYPE="CPT"',FALSE) ;
      If QSouche.Eof Then
        BEGIN
        QSouche.Insert ; InitNew(QSouche) ;
        QSouche.FindField('SH_TYPE').AsString:='CPT' ;
        QSouche.FindField('SH_SOUCHE').AsString:=QJalJ_JOURNAL.AsString ;
        QSouche.FindField('SH_LIBELLE').AsString:='Souche '+QJalJ_JOURNAL.AsString ;
        QSouche.FindField('SH_ABREGE').AsString:=QSouche.FindField('SH_LIBELLE').AsString ;
        QSouche.FindField('SH_NUMDEPART').AsInteger:=1 ;
        QSouche.FindField('SH_NUMDEPARTS').AsInteger:=1 ;
        QSouche.FindField('SH_NUMDEPARTP').AsInteger:=1 ;
        QSouche.FindField('SH_SOUCHEEXO').AsString:='-' ;
        QSouche.FindField('SH_DATEDEBUT').AsDateTime:=EncodeDate(1899,12,30) ;
        QSouche.FindField('SH_DATEFIN').AsDateTime:=EncodeDate(2099,12,31) ;
        If (NatJal='ODA') Or (NatJal='ANA') Then QSouche.FindField('SH_ANALYTIQUE').AsString:='X' Else QSouche.FindField('SH_ANALYTIQUE').AsString:='-' ;
        (*If Simu Then QSouche.FindField('SH_SIMULATION').AsString:='X' Else*) QSouche.FindField('SH_SIMULATION').AsString:='-' ;
        QSouche.Post ;
        Ferme(QSouche) ;
        AvertirTable('TTSOUCHECOMPTA') ;
        J_COMPTEURNORMAL.Reload ;
        QJalJ_COMPTEURNORMAL.AsString:=QJalJ_JOURNAL.AsString ;
        END Else
        BEGIN
        Ferme(QSouche) ;
        Pages.ActivePage:=PComplement ; J_COMPTEURNORMAL.SetFocus ;
        MsgBox.Execute(26,'','') ; Exit ;
        END ;
end
else
begin
        Pages.ActivePage:=PComplement ; J_COMPTEURNORMAL.SetFocus ;
        MsgBox.Execute(26,'','') ; Exit ;
end ;
      end;
   END ;
if VH^.CPIFDEFCEGID then
Begin
      If (QJalJ_COMPTEURSIMUL.AsString='') Then QJalJ_COMPTEURSIMUL.AsString:='SIM' ;
end else
begin
   { AJOUT ME LE 18/04/2001 }
   If (QJalJ_COMPTEURSIMUL.AsString='')  and  (ctxPCL in V_PGI.PGIContexte)
    and (NatJal<>'ODA') and (NatJal<>'ANA') and (NatJal<>'ANO') and (NatJal<>'CLO')Then
       if ExisteSQL ('SELECT * FROM SOUCHE WHERE SH_TYPE="CPT" AND SH_SOUCHE="SIM"') then
        QJalJ_COMPTEURSIMUL.AsString  := 'SIM';
end ;
(**
   If ((NatJal<>'ODA') and (NatJal<>'ANA') and (QJalJ_COMPTEURSIMUL.AsString='')) Then
      BEGIN
      Pages.ActivePage:=PComplement ; J_COMPTEURSIMUL.SetFocus ;
      MsgBox.Execute(26,'','') ; Exit ;
      END ;
**)
   END ;
if QJal.state in [dsInsert] then
   BEGIN
   if CodExist then BEGIN Pages.ActivePage:=PCar ; J_Journal.SetFocus ; MsgBox.Execute(4,'','') ; Exit ; END ;
   END ;
DateModification(QJal,'J') ; ClearChampsSurEnregOk ;
Result:=TRUE  ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/05/2002
Modifié le ... : 21/10/2002
Description .. : valeur par defaut pour la date à prendre en compte dans les 
Suite ........ : fichiers etebac
Suite ........ : 
Suite ........ : - 21/10/2002 - initialisation de control du panel saisie
Mots clefs ... : 
*****************************************************************}
Procedure TFJournal.NewEnreg ;
BEGIN
InitNew(QJal) ;
AvecMvt:=False ;
Pages.ActivePage:=PCar ;
QJalJ_NATUREJAL.Value:='OD' ;
QJalJ_MODESAISIE.Value:='-' ;
QJalJ_CHOIXDATE.Value:= 'DATEOPERATION' ; //LG*
J_NATUREJALChange(Nil) ;
DateCreation(QJal,'J') ;
J_Journal.Enabled:=TRUE ; J_Journal.SetFocus ;
J_INCREF.Checked:=false ; J_EQAUTO.Checked:=false ; J_NATCOMPL.Checked:=false ; J_INCNUM.Checked:=false ;
END ;

Function TFJournal.CodExist : Boolean ;
BEGIN Result:=Presence('JOURNAL','J_JOURNAL',QJalJ_JOURNAL.AsString) ; END ;

Function TFJournal.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert: if Not OnSauve  then Exit ;
   nbPost          : if Not EnregOK  then Exit ;
   nbDelete        : if Not Supprime then Exit ;
   end ;
if (QJal.State=dsInsert) And (Button=nbPost) then DevalideLexoSurCreatJal ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[18]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFJournal.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFJournal.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFJournal.BPriorClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFJournal.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFJournal.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFJournal.BinsertClick(Sender: TObject);
begin if ExJaiLeDroitConcept(TConcept(ccJalCreat),True) then Bouge(nbInsert) ; end;

procedure TFJournal.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then begin Bouge(nbInsert) ; exit ; end ;
   if (V_PGI.MonoFiche) or (Mode=taCreatOne) then Close ;
   END ;
end;

procedure TFJournal.BImprimerClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
PlanJournal(QJalJ_JOURNAL.AsString,True) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFJournal.SJalDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not QJal.Modified ;
BCumul.Enabled:=Not(QJal.State in [dsInsert]) ;
BZecrimvt.Enabled:=Not(QJal.State in [dsInsert]) ;
BMenuZoom.Enabled:=Not(QJal.State in [dsInsert]) ;
BImprimer.Enabled:=Not(QJal.State in [dsInsert]) ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QJal.BOF;
   DnEnable := Enabled and not QJal.EOF;
   BFirst.Enabled := UpEnable; BPrior.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
   if ((Field.FieldName='J_LIBELLE') and (QJalJ_ABREGE.AsString='')) then
      QJalJ_ABREGE.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFJournal.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFJournal.BFermeClick(Sender: TObject);
begin Close ; end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/05/2002
Modifié le ... : 21/10/2002
Description .. : Gestion de l'affichage du panel de saisie de tresorerie
Suite ........ :
Suite ........ : - 09/08/2002 - recherche de la nature de piece en fonction
Suite ........ : de la nature du journal
Suite ........ :
Suite ........ : - 21/10/2002 - le panel de saisie des info bordereau n'est
Suite ........ : visible que pour un bordereau
Mots clefs ... : 
*****************************************************************}
Procedure TFJournal.AutoriseOuNonLesChamps ;
Var NatJal : String ;
BEGIN
NatJal:=J_NATUREJAL.Value ;
J_AXE.Enabled:=(NatJal='ODA') Or (NatJal='ANA') ; TJ_AXE.Enabled:=(NatJal='ODA') Or (NatJal='ANA') ;
J_Axe.Color := IIF(J_Axe.Enabled, ClWindow, ClBtnFace);

if(NatJal='ANO') Or (NatJal='CLO') then J_CENTRALISABLE.Enabled:=False
                                   else J_CENTRALISABLE.Enabled:=True ;

if(NatJal='ANO') Or (NatJal='CLO') Or (NatJal='ODA') Or (NatJal='ANA') then
  BEGIN J_COMPTEINTERDIT.Enabled:=False ; TJ_COMPTEINTERDIT.Enabled:=False ; END else
  BEGIN J_COMPTEINTERDIT.Enabled:=True  ; TJ_COMPTEINTERDIT.Enabled:=True  ; END ;
J_CompteInterdit.Color := IIF(J_CompteInterdit.Enabled, ClWindow, ClBtnFace);

if(NatJal='ANO') Or (NatJal='BQE') Or
  (NatJal='CAI') Or (NatJal='ODA') Or
  (NatJal='CLO') Or (NatJal='ANA') then
  BEGIN J_COMPTEAUTOMAT.Enabled:=False ; TJ_COMPTEAUTOMAT.Enabled:=False ; END else
  BEGIN J_COMPTEAUTOMAT.Enabled:=True  ; TJ_COMPTEAUTOMAT.Enabled:=True  ; END ;
J_CompteAutomat.Color := IIF(J_CompteAutomat.Enabled, ClWindow, ClBtnFace);

J_COMPTEURSIMUL.Enabled:=(NatJal<>'ODA') And (NatJal<>'ANA') And (NatJal<>'ANO') ;
J_CompteursIMul.Color := IIF(J_CompteursIMul.Enabled, ClWindow, ClBtnFace);
TJ_COMPTEURSIMUL.Enabled:=(NatJal<>'ODA') And (NatJal<>'ANA') And (NatJal<>'ANO') ;

J_CONTREPARTIE.Enabled :=((NatJal='BQE') Or (NatJal='CAI') or (QJalJ_EFFET.AsString='X')) ;
TJ_CONTREPARTIE.Enabled:=J_CONTREPARTIE.Enabled ;
J_Contrepartie.Color := IIF(J_ContrePartie.Enabled, ClWindow, ClBtnFace);

J_TYPECONTREPARTIE.Enabled :=((NatJal='BQE') Or (NatJal='CAI') Or (QJalJ_EFFET.AsString='X')) ;
J_TypeContrePartie.Color   := IIF(J_TypeContrePartie.Enabled, ClWindow, ClBtnFace);
TJ_TYPECONTREPARTIE.Enabled:=((NatJal='BQE') Or (NatJal='CAI') Or (QJalJ_EFFET.AsString='X')) ;

if (NatJal='BQE') then J_CONTREPARTIE.ZoomTable:=tzGbanque ;
if (NatJal='CAI') then J_CONTREPARTIE.ZoomTable:=tzGcaisse ;
if QJalJ_EFFET.AsString='X' then J_CONTREPARTIE.ZoomTable:=tzGEffet ;

J_NATUREJAL.Enabled:=(Not AvecMvt) ;
J_EFFET.Enabled:=(NatJal='OD') ;
J_MODESAISIE.Enabled:=(Not AvecMvt) and (NatJal<>'ANA') and (NatJal<>'ODA') and (NatJal<>'ANO') and (NatJal<>'CLO') and  (NatJal<>'ECC') ;
// GCO - 19-04-2002
J_ModeSaisie.Color  := IIF(J_MODESAISIE.Enabled, ClWindow, ClBtnFace);
if (QJal.State=dsInsert) then
  if not J_ModeSaisie.Enabled then begin
    J_ModeSaisie.ItemIndex := 0;
    QJalJ_MODESAISIE.Value:='-';
  end;
// FIN GCO

If (NatJal<>'CLO') And (NatJal<>'ANO') Then
   BEGIN
   J_COMPTEURNORMAL.Enabled:=(Not AvecMvt) ; TJ_COMPTEURNORMAL.Enabled:=(Not AvecMvt) ;
   J_CompteurNormal.Color := IIF(J_CompteurNormal.Enabled, ClWindow, ClBtnFace);
   END ;
if(J_MULTIDEVISE.Checked)And(AvecMvt)then J_MULTIDEVISE.Enabled:=False
                                     else J_MULTIDEVISE.Enabled:=True ;
if NatJal='CLO' then J_MULTIDEVISE.Enabled:=False ;
if(NatJal='ODA')Or(NatJal='ANA') then
  BEGIN
  J_MULTIDEVISE.Enabled:=False ;
  J_COMPTEURNORMAL.DataType:='ttSoucheComptaODA' ;
  END else J_COMPTEURNORMAL.DataType:='ttSoucheCompta' ;
PReg.TabVisible:=(NatJal='REG') ;
(*P.FUGIER 05/00 : DEBUT*)
InitFolio ;
(*P.FUGIER 05/00 : FIN*)
//LG* gestion de l'affichage du panel tresorerie
PTreso.TabVisible := ( J_NATUREJAL.Value = 'BQE' ) and (ctxPCL in V_PGI.PGIContexte);
PSaisie.TabVisible:= (( J_MODESAISIE.Value = 'BOR' ) or ( J_MODESAISIE.Value = 'LIB' )) ;
// Utiliser la bonne nature de pièce en fonction de la nature du journal
case CaseNatJal(NatJal) of
   tzJVente       : J_NATDEFAUT.DataType:='ttNatPieceVente' ;
   tzJAchat       : J_NATDEFAUT.DataType:='ttNatPieceAchat' ;
   tzJBanque      : J_NATDEFAUT.DataType:='ttNatPieceBanque' ;
   tzJEcartChange : J_NATDEFAUT.DataType:='ttNatPieceEcartChange' ;
   tzJOD          : J_NATDEFAUT.DataType:='ttNaturePiece' ;
   end ;
END ;

procedure TFJournal.J_NATUREJALChange(Sender: TObject);
begin
  if (CtxPcl in V_Pgi.PgiContexte) and (Trim(J_Journal.Text) = '') then Exit;
  AutoriseOuNonLesChamps ;
end;

Function TFJournal.ExisteLeCompte(LeWhere : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select G_GENERAL From GENERAUX '+LeWhere,True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

Function TFJournal.FaitLeWhere(St : String) : String ;
Var S1,S2 : String ;
BEGIN
if Pos(':',St)>0 then
   BEGIN
   S1:=Copy(St,1,Pos(':',St)-1) ; S2:=Copy(St,Pos(':',St)+1,Length(St)) ;
   S2:=BourreLaDonc(S2,fbGene) ;
   Result:='Where G_GENERAL >="'+S1+'%" And G_GENERAL <="'+S2+'"' ;
   END else
   BEGIN
   if Length(St)<VH^.Cpta[fbGene].Lg then Result:='Where G_GENERAL Like"'+St+'%"'
                                     else Result:='Where G_GENERAL ="'+St+'"' ;
   END ;
END ;

Function TFJournal.ChaineCpteValide(Sender : TObject) : Boolean ;
Var St,StCode,Stn : String ;
    LeWhere : String ;
BEGIN
Result:=True ;
if TCustomEdit(Sender).Text=''then Exit ;
St:=UpperCase(TCustomEdit(Sender).Text) ;
Stn:=TCustomEdit(Sender).Name ;
if St[Length(St)]<>';' then St:=St+';' ;
While St<>'' do
  BEGIN
  StCode:=ReadTokenSt(St) ;
  LeWhere:=FaitLeWhere(StCode) ;
  if Not ExisteLeCompte(LeWhere) then
     BEGIN
     if(Stn='J_COMPTEINTERDIT')Or(Stn='J_COMPTEAUTOMAT')
       Or(Stn='J_CONTREPARTIE') then Pages.ActivePage:=PComplement
                                else Pages.ActivePage:=PReg ;
     TCustomEdit(Sender).SetFocus ; MsgBox.Execute(6,'','') ;
     Result:=False ; Exit ;
     END ;
  END ;
END ;

Function TFJournal.DansUnCpte(Cpte1,Cpte2 : TObject) : Boolean ;
Var MemoCpte2,St,St1 : String ;
    S1,S2,LeWhere1,LeWhere2 : String ;
BEGIN
Result:=False ;
if TCustomEdit(Cpte1).Text='' then Exit ;
if TCustomEdit(Cpte2).Text='' then Exit ;
S1:=TCustomEdit(Cpte1).Text ; S2:=TCustomEdit(Cpte2).Text ;
if S1[Length(S1)]<>';' then S1:=S1+';' ;
if S2[Length(S2)]<>';' then S2:=S2+';' ;
MemoCpte2:=S2 ;
While S1<>'' do
  BEGIN
  St:=ReadTokenSt(S1) ;
  LeWhere1:=FaitLeWhere(St) ;
  Q1.Sql.Clear ; Q1.Sql.Add('Select G_GENERAL From GENERAUX '+LeWhere1);
  ChangeSQL(Q1) ; Q1.Open ;
  While Not Q1.EOF do
    BEGIN
    S2:=MemoCpte2 ;
    While S2<>'' do
       BEGIN
       St1:=ReadTokenSt(S2) ;
       LeWhere2:=FaitLeWhere(St1) ;
       Q2.Sql.Clear ; Q2.Sql.Add('Select G_GENERAL From GENERAUX '+LeWhere2);
       ChangeSQL(Q2) ; Q2.Open ;
       While Not Q2.Eof do
         BEGIN
         if Q1.Fields[0].AsString=Q2.Fields[0].AsString then
             BEGIN
             Result:=True ; Q1.Close ; Q2.Close ; Exit ;
             END ;
         Q2.Next ;
         END ;
       END ;
    Q1.Next ; Q2.Close ;
    END ;
  Q1.Close ;
  END ;
END ;

Function TFJournal.ContrePartieValide : Boolean;
Var St : String ;
BEGIN
Result:=False ;
St:=QJalJ_CONTREPARTIE.AsString ;
if St='' then
   BEGIN
   Pages.ActivePage:=PComplement ; J_CONTREPARTIE.SetFocus ;
   MsgBox.Execute(9,'','') ; Exit ;
   END ;
if QJalJ_TYPECONTREPARTIE.AsString='' then
   BEGIN
   Pages.ActivePage:=PComplement ; J_TYPECONTREPARTIE.SetFocus ;
   MsgBox.Execute(10,'','') ; Exit ;
   END ;
if Not CpteDeRegulValide(J_CONTREPARTIE) then
   BEGIN
   Pages.ActivePage:=PComplement ; J_CONTREPARTIE.SetFocus ;
   MsgBox.Execute(25,'','') ; Exit ;
   END ;
if DansUnCpte(J_COMPTEINTERDIT,J_CONTREPARTIE) then
   BEGIN
   Pages.ActivePage:=PComplement ; J_CONTREPARTIE.SetFocus ;
   MsgBox.Execute(20,'','') ; Exit ;
   END ;
Result:=True ;
END ;

Function TFJournal.ContrePartieDevise : Boolean;
Var QLoc : TQuery ;
    St,StDev : String ;
BEGIN
Result:=True ;
if QJalJ_EFFET.AsString='X' then Exit ;
St:=QJalJ_CONTREPARTIE.AsString ; StDev:='' ;
if St='' then Exit ;
QLoc:=OpenSql('Select BQ_DEVISE From BANQUECP Where BQ_GENERAL="'+St+'"',True) ;
if QLoc.Eof then BEGIN Ferme(QLoc) ; Exit ; END ;
StDev:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
if(StDev='')Or(StDev=V_PGI.DevisePivot)then Exit ;
if QJalJ_MULTIDEVISE.AsString='X' then Exit ;
MsgBox.Execute(24,'','') ; Pages.ActivePage:=Pcar ; J_MULTIDEVISE.SetFocus ; Result:=False ;
END ;

Function TFJournal.CpteDeRegulValide(Sender : TObject) : Boolean ;
BEGIN
Result:=True ;
if THDBCpteEdit(Sender).Text='' then Exit ;
if THDBCpteEdit(Sender).ExisteH=0 then Result:=False ;
END ;

Procedure TFJournal.MsgCpteRegul(Sender : TOBject ; i : Byte) ;
BEGIN
Pages.ActivePage:=PReg ; THDBCpteEdit(Sender).SetFocus ;
MsgBox.Execute(i,'','') ;
END ;

Function TFJournal.CrtlCpteDeRegul : Boolean ;
Var i : Byte ;
BEGIN
Result:=False ;
for i:=1 to 3 do
   BEGIN
    if Not CpteDeRegulValide(THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+'')))then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+'')),21) ;
       Exit ;
       END ;
    if DansUnCpte(J_COMPTEAUTOMAT,THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+''))) then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+'')),22) ;
       Exit ;
       END ;
    if DansUnCpte(J_COMPTEINTERDIT,THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+''))) then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULDEBIT'+IntToStr(i)+'')),23) ;
       Exit ;
       END ;
    if Not CpteDeRegulValide(THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+'')))then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+'')),21) ;
       Exit ;
       END ;
    if DansUnCpte(J_COMPTEAUTOMAT,THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+''))) then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+'')),22) ;
       Exit ;
       END ;
    if DansUnCpte(J_COMPTEINTERDIT,THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+''))) then
       BEGIN
       MsgCpteRegul(THDBCpteEdit(FindComponent('J_CPTEREGULCREDIT'+IntToStr(i)+'')),23) ;
       Exit ;
       END ;
   END ;
Result:=True ;
END ;

Function TFJournal.CaractereValide(Sender : TObject) : Boolean ;
Var i : Byte ;
    St,Stn : String ;
BEGIN
Result:=True ;
St:=TCustomEdit(Sender).Text ;
Stn:=TCustomEdit(Sender).Name ;
if St=''then Exit ;
St:=UpperCase(St) ;
for i:=1 to Length(St) do
   BEGIN
   if Not(St[i]in ['0'..'9','A'..'Z',';',':',VH^.Cpta[fbGene].Cb]) then
      BEGIN
      if(Stn='J_COMPTEINTERDIT')Or(Stn='J_COMPTEAUTOMAT')
        Or(Stn='J_CONTREPARTIE') then Pages.ActivePage:=PComplement
                                 else Pages.ActivePage:=PReg ;
      TCustomEdit(Sender).SetFocus ; MsgBox.Execute(5,'','') ;
      Result:=False ; Exit ;
      END ;
   END ;
END ;

Function TFJournal.EstJalAnal : Boolean ;
BEGIN
Result:=True ;
if(J_NATUREJAL.Value='ODA')Or(J_NATUREJAL.Value='ANA') then
   BEGIN
   if J_AXE.Value='' then
      BEGIN
      Pages.ActivePage:=PCar ; J_AXE.SetFocus ;
      MsgBox.Execute(19,'','') ; Result:=False ;
      END ;
   END
END ;

procedure TFJournal.J_CPTEREGULDEBIT1Exit(Sender: TObject);
begin if Mode<>taConsult then THDBCpteEdit(Sender).ExisteH ; end;

procedure TFJournal.BCumulClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
CumulCpteMensuel(fbJal,QJalJ_JOURNAL.AsString,QJalJ_LIBELLE.AsString,VH^.Entree) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFJournal.BZecrimvtClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
if(J_NATUREJAL.Value='ODA') Or (J_NATUREJAL.Value='ANA') then
   VisuZoomAna(J_JOURNAL.Text,'','',J_AXE.Value,EXRF(VH^.Entree.Code)) else
   ZoomEcritureMvt(J_JOURNAL.Text, fbJal,'MULMMVTS' ) ;
{$ENDIF}
{$ENDIF}
end;

// Pour la suppression ======*************---------------
Function TFJournal.Supprime : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,'','')<>mrYes then Exit ;
Result:=True ;
END ;

Function EstDansAnalytiq(St : String) : Boolean ;
BEGIN Result:=Presence('ANALYTIQ','Y_JOURNAL',St) ; END ;

Function EstDansEcriture(St : String) : Boolean ;
BEGIN Result:=Presence('ECRITURE','E_JOURNAL',St) ; END ;

Function EstDansSociete(St : String) : Boolean ;
Var
{$IFDEF SPEC302}
QLoc : TQuery ;
{$ENDIF}
    StLoc : String ;
BEGIN
{$IFDEF SPEC302}
QLoc:=OpenSql('Select SO_JALFERME,SO_JALOUVRE From SOCIETE Where SO_JALFERME="'+St+'" OR '+
              'SO_JALOUVRE="'+St+'"',True) ;
Result:=(Not QLoc.EOF) ;
Ferme(QLoc) ;
{$ELSE}
Result:=True ;
StLoc:=GetParamsoc('SO_JALOUVRE') ; if StLoc=St then Exit ;
StLoc:=GetParamsoc('SO_JALFERME') ; if StLoc=St then Exit ;
Result:=False ;
{$ENDIF}
END ;

Function  EstDansPiece(St : String)   : Boolean ;    // Modif Fiche 10856 SBO
BEGIN Result:=Presence('PIECE','GP_JALCOMPTABLE',St) ; END ;

Function  EstDansParamPiece(St : String)   : Boolean ;    // Modif Fiche 10856 SBO
BEGIN Result:=Presence('PARPIECE','GPP_JOURNALCPTA',St) ; END ;

Function  EstDansParamPieceCompl(St : String)   : Boolean ;    // Modif Fiche 10856 SBO
BEGIN Result:=Presence('PARPIECECOMPL','GPC_JOURNALCPTA',St) ; END ;

Function EstDansSouche(St : String) : Boolean ;
BEGIN Result:=Presence('SOUCHE','SH_JOURNAL',St) ; END ;

procedure TFJournal.QJalAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFJournal.QJalAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFJournal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then AvertirMultiTable('ttJournal') ;
END ;

procedure TFJournal.BMenuZoomClick(Sender: TObject);
BEGIN PopZoom(BMenuZoom,POPZ) ; end;

procedure TFJournal.BJALDIVClick(Sender: TObject);
(*Var Crit : TCritJour ;
    D1,D2 : TdateTime ;*)
begin
(*Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=V_PGI.Encours.Deb ; D2:=V_PGI.Encours.Fin ;
If V_PGI.Entree.Code=V_PGI.Suivant.Code Then BEGIN D1:=V_PGI.Suivant.Deb ; D2:=V_PGI.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
InitCritJour(Crit) ;
Crit.Code1:=J_Journal.text ;
Crit.Code2:=Crit.Code1 ;
JalDivisioZoom(Crit) ; *)
end ;

procedure TFJournal.BJALCPTClick(Sender: TObject);
(*Var Crit : TCritJour ;
    D1,D2 : TdateTime ;*)
begin
(*Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=V_PGI.Encours.Deb ; D2:=V_PGI.Encours.Fin ;
If V_PGI.Entree.Code=V_PGI.Suivant.Code Then BEGIN D1:=V_PGI.Suivant.Deb ; D2:=V_PGI.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
InitCritJour(Crit) ;
Crit.Code1:=J_Journal.text ;
Crit.Code2:=Crit.Code1 ;
JalCpteGeZoom(Crit) ;*)
end;

procedure TFJournal.BJALPERClick(Sender: TObject);
(*Var Crit : TCritJourGene ;
    D1,D2 : TdateTime ;*)
begin
(*Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=V_PGI.Encours.Deb ; D2:=V_PGI.Encours.Fin ;
If V_PGI.Entree.Code=V_PGI.Suivant.Code Then BEGIN D1:=V_PGI.Suivant.Deb ; D2:=V_PGI.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
InitCritJourGene(Crit) ;
Crit.Code1:=J_Journal.text ;
Crit.Code2:=Crit.Code1 ;
JalPeriodeZoom(Crit) ;*)
end;

procedure TFJournal.BJALCENTClick(Sender: TObject);
(*Var Crit : TCritJourGene ;
    D1,D2 : TdateTime ;*)
begin
(*Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=V_PGI.Encours.Deb ; D2:=V_PGI.Encours.Fin ;
If V_PGI.Entree.Code=V_PGI.Suivant.Code Then BEGIN D1:=V_PGI.Suivant.Deb ; D2:=V_PGI.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
InitCritJourGene(Crit) ;
Crit.Code1:=J_Journal.text ;
Crit.Code2:=Crit.Code1 ;
JalCentralZoom(Crit) ; *)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 23/07/2002
Modifié le ... : 23/09/2002
Description .. : - 23/07/2002 - creation des champs de la requete 
Suite ........ : dynamiquement en fct de la version
Suite ........ : -23/09/2002- suppression du champs j_natdef
Mots clefs ... : 
*****************************************************************}
procedure TFJournal.FormCreate(Sender: TObject);
begin
{$IFNDEF IMP}
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
{$ENDIF}
Q:=NIL ; LeCompte:='aze' ;
QJalJ_INCREF:= TStringField.Create(QJal) ;
QJalJ_INCREF.FieldName:='J_INCREF' ;
QJalJ_INCREF.FieldKind:=fkData ;
QJalJ_INCREF.DataSet:=QJal;
QJalJ_NATCOMPL:= TStringField.Create(QJal) ;
QJalJ_NATCOMPL.FieldName:='J_NATCOMPL' ;
QJalJ_NATCOMPL.FieldKind:=fkData ;
QJalJ_NATCOMPL.DataSet:=QJal;
QJalJ_INCNUM:= TStringField.Create(QJal) ;
QJalJ_INCNUM.FieldName:='J_INCNUM' ;
QJalJ_INCNUM.FieldKind:=fkData ;
QJalJ_INCNUM.DataSet:=QJal;
QJalJ_EQAUTO:= TStringField.Create(QJal) ;
QJalJ_EQAUTO.FieldName:='J_EQAUTO' ;
QJalJ_EQAUTO.FieldKind:=fkData ;
QJalJ_EQAUTO.DataSet:=QJal;
QJalJ_NATDEFAUT:= TStringField.Create(QJal) ;
QJalJ_NATDEFAUT.FieldName:='J_NATDEFAUT' ;
QJalJ_NATDEFAUT.FieldKind:=fkData ;
QJalJ_NATDEFAUT.DataSet:=QJal;
end;

procedure TFJournal.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F3  : BPriorClick(Nil) ;
        VK_F4  : BNextClick(Nil) ;
        VK_F10 : BValiderClick(Nil) ;
        VK_RETURN :BEGIN
                   if ActiveControl is TCustomMemo then Exit ;
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

Procedure TFJournal.DevalideLexoSurCreatJal ;
Var i : Byte ;
    St : String ;
BEGIN
St:='' ; for i:=1 to 24 do St:=St+'-' ;
ExecuteSql('Update EXERCICE Set EX_VALIDEE="'+St+'" Where EX_ETATCPTA="OUV"') ;
END ;

Procedure TFJournal.FormateLesMontants ;
BEGIN
QJalJ_DEBITDERNMVT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_CREDITDERNMVT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTALDEBIT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTALCREDIT.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTDEBP.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTCREP.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTDEBE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTCREE.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTDEBS.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QJalJ_TOTCRES.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
ChangeMask(JSOLCREP,V_PGI.OkdecV,'') ;
ChangeMask(JSOLCREE,V_PGI.OkdecV,'') ;
ChangeMask(JSOLCRES,V_PGI.OkdecV,'') ;
END ;

procedure TFJournal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFJournal.J_EFFETClick(Sender: TObject);
begin
FAvertir:=TRUE ;
J_CONTREPARTIE.Enabled :=((J_NATUREJAL.Value='BQE') Or (J_NATUREJAL.Value='CAI') or (J_EFFET.Checked)) ;
TJ_CONTREPARTIE.Enabled:=J_CONTREPARTIE.Enabled ;
J_TYPECONTREPARTIE.Enabled :=((J_NATUREJAL.Value='BQE') Or (J_NATUREJAL.Value='CAI') Or (J_EFFET.Checked)) ;
TJ_TYPECONTREPARTIE.Enabled:=J_TYPECONTREPARTIE.Enabled ;
if (J_NATUREJAL.Value='BQE') then J_CONTREPARTIE.ZoomTable:=tzGbanque ;
if (J_NATUREJAL.Value='CAI') then J_CONTREPARTIE.ZoomTable:=tzGcaisse ;
if J_EFFET.Checked then J_CONTREPARTIE.ZoomTable:=tzGEffet ;
end;

(*P.FUGIER 05/00 : DEBUT*)
procedure TFJournal.J_MODESAISIEChange(Sender: TObject);
begin
InitFolio ;
end ;
(*P.FUGIER 05/00 : FIN*)

procedure TFJournal.J_JOURNALExit(Sender: TObject);
begin
if QJal.State=dsInsert then QJALJ_JOURNAL.AsString:=Trim(QJALJ_JOURNAL.AsString) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2002
Modifié le ... : 18/04/2002
Description .. : Active ou Désactive les contrôles de la fiche lors de
Suite ........ : la création d'un compte auxiliaire afin de forcer
Suite ........ : la saisie de  J_JOURNAL et J_NATUREJAL
Mots clefs ... :
*****************************************************************}
procedure TFJournal.ActivationControl;
var lStJournal : string;
begin
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  lStJournal := Trim( J_Journal.Text );

  J_Axe.Enabled := LStJournal <> '';
  J_Axe.Color   := IIF( J_Axe.Enabled, ClWindow, ClBtnFace );

  J_Libelle.Enabled := LStJournal <> '';
  J_Libelle.Color   := IIF( J_Libelle.Enabled, ClWindow, ClBtnFace );

  J_Abrege.Enabled := LStJournal <> '';
  J_Abrege.Color   := IIF( J_Abrege.Enabled, ClWindow, ClBtnFace );

  J_ModeSaisie.Enabled := LStJournal <> '';
  J_ModeSaisie.Color   := IIF(J_ModeSaisie.Enabled, ClWindow, ClBtnFace );

  J_MultiDevise.Enabled   := LStJournal <> '';
  J_Centralisable.Enabled := LStJournal <> '';
  J_Effet.Enabled         := LStJournal <> '';

  if lStJournal <> '' then
    AutoriseOuNonLesChamps;

end;

procedure TFJournal.PagesChanging(Sender: TObject; var AllowChange: Boolean);
begin
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  AllowChange := Trim(J_Journal.Text) <> '';

  if not Allowchange then
    J_Journal.SetFocus;
end;

procedure TFJournal.J_JOURNALChange(Sender: TObject);
begin
  ActivationControl;
end;

end.
