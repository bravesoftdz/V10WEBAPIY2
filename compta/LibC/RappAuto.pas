unit RappAuto;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Spin, StdCtrls, Hctrls, Mask, Buttons, ExtCtrls,
  //ULibEcriture,
  LetRegul,
  UTOB,   
  UtilSoc,
{$IFNDEF SANSCOMPTA}
  Hcompte,
{$ENDIF}
  hmsgbox, LettAuto,
{$IFNDEF IMP}
{$IFNDEF GCGC}
  {$IFDEF SPEC302} Societe,{$ENDIF}
{$ENDIF}
{$ENDIF}
 {$IFNDEF EAGLCLIENT}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  HQry,
 {$ENDIF}
  ComCtrls, HStatus, Ent1, LettUtil,  SaisUtil,
  HSysMenu, SaisComm, ed_tools, HPanel, UiUtil, HTB97, ParamDat, Paramsoc, UtilPgi
{$IFNDEF GCGC}
  ,DelVisuE
{$ENDIF}
   ;

procedure RapprochementAuto ( LeGene,LeAux : String ) ;
procedure RapprochementAutoMP ( LeGene,LeAux : String ; Qui : tProfilTraitement) ;


type
  TFRappAuto = class(TForm)
    HPB: TToolWindow97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    PParam: TPanel;
    GParam       : TGroupBox;
    CMontant     : TCheckBox;
    CSolde       : TCheckBox;
    CRef         : TCheckBox;
    MAXPROF      : TSpinEdit;
    H_MaxProf    : THLabel;
    H_MaxDuree   : THLabel;
    MAXDUREE     : TSpinEdit;
    HRappro      : THMsgBox;
    BStop: TToolbarButton97;
    BGenere: TToolbarButton97;
    BParam: TToolbarButton97;
    HMTrad: THSystemMenu;
    Pages: TPageControl;
    PCrit: TTabSheet;
    HLabel4: THLabel;
    E_GENERAL: THCpteEdit;
    HLabel5: THLabel;
    E_GENERAL_: THCpteEdit;
    Label14: THLabel;
    T_NATUREAUXI: THValComboBox;
    HLabel1: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    HLabel2: THLabel;
    E_AUXILIAIRE_: THCpteEdit;
    LETTREPARTIEL: TCheckBox;
    Bevel1: TBevel;
    PLibres: TTabSheet;
    Bevel2: TBevel;
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
    Dock: TDock97;
    CDEV: THValComboBox;
    CAvecOD: TCheckBox;
    CODSALFOU: TCheckBox;
    TSEcritures: TTabSheet;
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    HLabel3: THLabel;
    E_REFLETTRAGE: TEdit;
    CAnc: TCheckBox;
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BStopClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BGenereClick(Sender: TObject);
    procedure BParamClick(Sender: TObject);
    procedure CMontantClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure CRefClick(Sender: TObject);
    procedure CAvecODClick(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    TBase,TLETT : TList ;
    GeneNumSol,CurBase,GeneInd  : integer ;
    Arret,OkTraite : Boolean ;
    IFL     : RINFOSLETT ;
    DEVI    : RDEVISE ;
   // ModeS   : String3 ;
    Qui : tProfilTraitement ;
    procedure ErgoS3S5 ;
{Recherches et chargements}
    procedure ChargeListeBase( vQ : TQuery ) ;
    procedure ChargeEcrCompte ( T : T_RAPPAUTO ) ;
    function  RechercheComptes : boolean ;
    function  ArretRappro : boolean ;
    procedure ChargeDevises ;
    procedure RecupInfosDevise ;
    function  GetCritLibres : String ;
    function  GetCritReste : String ;
{Lettrage REFERENCE}
    procedure LettrerReference ;
    function  QuelleRef ( L : TL_Rappro ) : String ;
    function  QuelleDate ( L : TL_Rappro ; Var ControleDate : boolean ) : TDateTime ;
    function  MemeRef ( LaRef,RefTest : String ; Client : boolean ) : Boolean ;
    function  BonneDateRef ( LaDate,DateTest : TDateTime ; Client : boolean ) : Boolean ;
    procedure SolutionsRef ( NumSol : integer ) ;
    procedure AnnuleSolutionRef ( NoSol : integer )  ;
    Function  BonneNature ( Nat1,Nat2 : String ) : boolean ;
{Lettrage MONTANT COMBINATOIRE}
    procedure LettrerMontant ;
    function  ConstruitListes ( icur : integer ; Debit : boolean ; Var LM : T_D ; Var LP,LG : T_I ) : integer ;
{Lettrage SOLDE NUL}
    procedure LettrerSolde ;
{Lettrage fichier}
    procedure LettrerUnCompte ;
    procedure FichierLettrage ( NoSol : integer ; Total : boolean ; ModeRappro : Char ) ;
    procedure FinirLettrage ( NoSol : integer ; Total : boolean ; DMin,DMax : TDateTime ; CodeL : String ; ModeRappro : Char ) ;
    procedure DelettrePartiel ( X : TL_Rappro ) ;
    Function  isISF ( L : TL_Rappro ) : Tisf ;
    procedure LettrerAnc ;
  public
    LeGene,LeAux : String;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  HEnt1;

procedure RapprochementAuto ( LeGene,LeAux : String ) ;
Var X  : TFRappAuto ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage','nrEnca','nrDeca','nrRelance'],True,'nrCloture') then Exit ;
X:=TFRappAuto.Create(Application) ;
X.LeGene:=LeGene ; X.LeAux:=LeAux ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Qui:=prAucun ;
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrCloture',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure RapprochementAutoMP ( LeGene,LeAux : String ; Qui : tProfilTraitement) ;
Var X  : TFRappAuto ;
    PP : THPanel ;
BEGIN
If Not MonProfilOk(Qui) Then
  BEGIN
  if _Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage','nrEnca','nrDeca','nrRelance'],True,'nrCloture') then Exit ;
  END ;
X:=TFRappAuto.Create(Application) ;
X.LeGene:=LeGene ; X.LeAux:=LeAux ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Qui:=Qui ;
     X.ShowModal ;
    Finally
     X.Free ;
     If Not MonProfilOk(Qui) Then _Bloqueur('nrCloture',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFRappAuto.BAbandonClick(Sender: TObject);
begin
 //if IsInside(Self) then exit ;
 if not OkTraite then
 begin
   Close;
   //SG6 08.02.05 FQ 15328
   if IsInside(Self) then
   begin
     CloseInsidePanel(Self) ;
   end;
 end
 else Arret:=True ;
end;

procedure TFRappAuto.BStopClick(Sender: TObject);
begin Arret:=True ; end;

procedure TFRappAuto.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VideListe(TBase) ; TBase.Free ; VideListe(TLett) ; TLett.Free ;
if Parent is THPanel then
   BEGIN
  If Not MonProfilOk(Qui) Then _Bloqueur('nrCloture',False) ;
   Action:=caFree ;
   END ;
end;

procedure TFRappAuto.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TBase:=TList.Create ; TLett:=TList.Create ;
end;

procedure TFRappAuto.ErgoS3S5 ;
BEGIN
{$IFDEF CCS3}
CODSALFOU.Visible:=False;
MaxDuree.Visible:=False ; H_MaxDuree.Visible:=False ;
MaxProf.MaxValue:=5 ;
LettrePartiel.Checked:=True ; LettrePartiel.Visible:=False ; 
{$ENDIF}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/03/2002
Modifié le ... : 06/12/2002
Description .. : LG* ChargeInfosLett deplace dans LettUtil
Suite ........ :  - 06/12/2002 - fiche 10763 - recuperation du parametre 
Suite ........ : utilisation des OD
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.FormShow(Sender: TObject);
begin
InitTablesLibresTiers(PLibres) ;
IFL:=ChargeInfosLett ; ChargeDevises ;
//LG* utilisation des param generaux pour coher les cases
  LETTREPARTIEL.Checked := IFL.AvecLetPartiel ;
  CAvecOD.Checked       := IFL.CAvecOD ;
OkTraite:=False ;
E_DATECOMPTABLE.Text:=stDate1900 ; E_DATECOMPTABLE_.Text:=stDate2099 ;
if LeGene<>'' then BEGIN E_GENERAL.Text:=LeGene ; E_GENERAL_.Text:=LeGene ; END ;
if LeAux<>'' then BEGIN E_AUXILIAIRE.Text:=LeAux ; E_AUXILIAIRE_.Text:=LeAux ; END ;
PositionneEtabUser(E_ETABLISSEMENT, False); // 15086
ErgoS3S5 ;
end;

procedure TFRappAuto.BGenereClick(Sender: TObject);
begin
{$IFNDEF GCGC}
if ((Not OkTraite) and (TBase.Count>0)) then VisuPiecesGenere(TBase,EcrGen,6) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/03/2002
Modifié le ... : 02/08/2007
Description .. : LG* ChargeInfosLett deplace dans LettUtil
Suite ........ : BVE 02.08.07 : suppression des differents IFDEF lancement 
Suite ........ : de ParamSociété dans tous les cas.
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.BParamClick(Sender: TObject);
begin
//LG* 16/01/2002 modif de l'appel de la branche des parametres
ParamSociete(False,'','SCO_LETTRAGE','',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000) ;
end;

{===================================== CHARGEMENTS ========================================}
procedure TFRappAuto.ChargeDevises ;
Var Q : TQuery ;
    St : String ;
BEGIN
CDEV.Items.Clear ; CDEV.Values.Clear ;
Q:=OpenSQL('Select D_DEVISE, D_DECIMALE, D_QUOTITE from DEVISE',True) ;
While Not Q.EOF do
   BEGIN
   CDEV.Items.Add(Q.Fields[0].AsString) ;
   St:=IntToStr(Q.Fields[1].AsInteger)+';'+FloatToStr(Q.Fields[2].AsFloat) ;
   CDEV.Values.Add(St) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - FB 14451 - exclusion des ecritures de type 
Suite ........ : bordereau et multi echeance
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.ChargeEcrCompte ( T : T_RAPPAUTO ) ;
Var SQL,StV8,SReste : String ;
    L        : TL_Rappro ;
    QEcr   : TQuery ;
BEGIN
VideListe(TLett) ;
//QEcr:=TQuery.Create(Application) ; QEcr.DataBaseName:='SOC' ;
//QECR.Close ; QECR.SQL.Clear ;
if CAnc.Checked then
  SQL:='Select * From Ecriture Left Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE left join moderegl on T_MODEREGLE=MR_MODEREGLE ' +
  ' Where E_AUXILIAIRE="'+T.Auxiliaire+'" AND E_GENERAL="'+T.General+'"' 
  else
   SQL:='Select * from Ecriture Where E_AUXILIAIRE="'+T.Auxiliaire+'" AND E_GENERAL="'+T.General+'"' ;

if LettrePartiel.Checked then SQL:=SQL+' AND (E_ETATLETTRAGE="AL" or E_ETATLETTRAGE="PL")' else SQL:=SQL+' AND E_ETATLETTRAGE="AL"' ;
SQL:=SQL+' AND E_DEVISE="'+T.Devise+'"' ;
StV8:=LWhereV8 ; if StV8<>'' then SQL:=SQL+' AND '+StV8 ;
SQL:=SQL+' AND '+LWhereBase(True,False,False,False) ;
SReste:=GetCritReste ; if SReste<>'' then SQL:=SQL+SReste ;
if CAnc.Checked then
 begin
  SQL:=SQL+' AND E_MODESAISIE="-" and MR_NOMBREECHEANCE=1 and E_DEVISE="'+V_PGI.DevisePivot+'" '  ;
  if not (CAvecOD.Checked) then SQL:=SQL+' AND E_NATUREPIECE<>"OD" '  ;
  SQL:=SQL+' order by E_AUXILIAIRE, E_GENERAL,E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE '
 end
  else
   SQL:=SQL+LTri ;
QEcr:=OpenSQL(SQL,true) ;
While Not QECR.EOF do
   BEGIN
   L:=TL_Rappro.Create ;
   {Comptes et caractéristiques}
   L.General:=QECR.FindField('E_GENERAL').AsString ; L.Auxiliaire:=QECR.FindField('E_AUXILIAIRE').AsString ;
   L.DateC:=QECR.FindField('E_DATECOMPTABLE').AsDateTime ; L.DateE:=QECR.FindField('E_DATEECHEANCE').AsDateTime ; L.DateR:=QECR.FindField('E_DATEREFEXTERNE').AsDateTime ;
   L.RefI:=QECR.FindField('E_REFINTERNE').AsString ; L.RefL:=QECR.FindField('E_REFLIBRE').AsString ;
   L.RefE:=QECR.FindField('E_REFEXTERNE').AsString ; L.Lib:=QECR.FindField('E_LIBELLE').AsString ;
   L.Jal:=QECR.FindField('E_JOURNAL').AsString ; L.Numero:=QECR.FindField('E_NUMEROPIECE').AsInteger;
   L.NumLigne:=QECR.FindField('E_NUMLIGNE').AsInteger ; L.NumEche:=QECR.FindField('E_NUMECHE').AsInteger ;
   L.CodeL:=QECR.FindField('E_LETTRAGE').AsString ; L.CodeD:=QECR.FindField('E_DEVISE').AsString ;
   L.TauxDEV:=QECR.FindField('E_TAUXDEV').AsFloat ;
   L.Nature:=QEcr.FindField('E_NATUREPIECE').AsString ;
   L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
   L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
   L.EditeEtatTva:=(QEcr.FindField('E_EDITEETATTVA').AsString='X') ;
   L.Solution:=0 ; L.Exo:=QEcr.FindField('E_EXERCICE').AsString ;
   {Montants, Lettrage}
   L.DebDev:=QECR.FindField('E_DEBITDEV').AsFloat ; L.CredDev:=QECR.FindField('E_CREDITDEV').AsFloat ;
   L.Debit:=QECR.FindField('E_DEBIT').AsFloat ; L.Credit:=QECR.FindField('E_CREDIT').AsFloat ;
   if L.CodeD<>V_PGI.DevisePivot then
      BEGIN
      L.DebitCur:=QECR.FindField('E_DEBITDEV').AsFloat ; L.CreditCur:=QECR.FindField('E_CREDITDEV').AsFloat ;
      END else
      BEGIN
       L.DebitCur:=QECR.FindField('E_DEBIT').AsFloat ; L.CreditCur:=QECR.FindField('E_CREDIT').AsFloat ;
      END ;
   {Objet}
   TLETT.Add(L) ;
   if ((CMontant.Checked) and (TLETT.Count>=MaxDroite)) then Break ;
   if ((Not CMontant.Checked) and (TLETT.Count>=2000)) then Break ;
   QECR.Next ;
   END ;
Ferme(QEcr) ;
END ;

procedure TFRappAuto.ChargeListeBase( vQ : TQuery ) ;
Var T : T_RAPPAUTO ;
BEGIN
VideListe(TBase) ;
While Not vQ.EOF do
   BEGIN
   T:=T_RAPPAUTO.Create ; T.General:=vQ.FindField('E_GENERAL').asString ;
   T.Auxiliaire:=vQ.FindField('E_AUXILIAIRE').AsString ;
   T.Devise:=vQ.FindField('DEV').AsString ;
   TBase.Add(T) ;
   vQ.Next ;
   END ;
END ;

function TFRappAuto.GetCritLibres : String ;
Var St : String ;
    CC : THCpteEdit ;
    i  : integer ;
BEGIN
St:='' ;
if Plibres.TabVisible then for i:=0 to 9 do
   BEGIN
   CC:=THCpteEdit(FindComponent('T_TABLE'+IntToStr(i))) ; if CC=Nil then Continue ;
   if CC.Text<>'' then St:=St+' AND T_TABLE'+IntToStr(i)+'="'+CC.Text+'"' ;
   END ;
Result:=St ;
END ;

function TFRappAuto.GetCritReste : String ;
Var St : String ;
BEGIN
St:='' ;
if E_JOURNAL.Value<>'' then St:=St+' AND E_JOURNAL="'+E_JOURNAL.Value+'"' ;
if E_DEVISE.Value<>''  then St:=St+' AND E_DEVISE="'+E_DEVISE.Value+'"' ;
if E_ETABLISSEMENT.Value<>''  then St:=St+' AND E_ETABLISSEMENT="'+E_ETABLISSEMENT.Value+'"' ;
if E_REFINTERNE.Text<>''  then St:=St+' AND E_REFINTERNE LIKE "'+E_REFINTERNE.Text+'%"' ;
if E_EXERCICE.Value<>'' then St:=St+' AND E_EXERCICE="'+E_EXERCICE.Value+'" ' ;
St:=St+' AND E_DATECOMPTABLE>="'+UsDate(E_DATECOMPTABLE)+'" AND E_DATECOMPTABLE<="'+UsDate(E_DATECOMPTABLE_)+'" ' ;
Result:=St ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/03/2002
Modifié le ... : 13/04/2004
Description .. : LG* suprression warning
Suite ........ : - LG - 13/04/2004 - ajout du lettrage par anc.
Mots clefs ... : 
*****************************************************************}
function TFRappAuto.RechercheComptes : boolean ;
Var
 SQL,Slibres,SReste,St : String ;
 lQ : TQuery ;
BEGIN
Result:=True ;
if CAnc.Checked then
 begin
  SQL:='Select E_GENERAL, E_AUXILIAIRE, Sum(E_DEBIT-E_CREDIT), Sum(E_DEBITDEV-E_CREDITDEV), Max(E_DEVISE) DEV '
    +'From Ecriture Left Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE left join moderegl on T_MODEREGLE=MR_MODEREGLE ' ;
  If Qui<>prAucun Then SQL:=SQL+' Left Join GENERAUX On E_GENERAL=G_GENERAL ' ;
 end
  else
   begin
    SQL:='Select E_GENERAL, E_AUXILIAIRE, Sum(E_DEBIT-E_CREDIT), Sum(E_DEBITDEV-E_CREDITDEV), Max(E_DEVISE) DEV '
         +'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;
    If Qui<>prAucun Then SQL:=SQL+' Left Outer Join GENERAUX On E_GENERAL=G_GENERAL ' ;
   end ;
SQL:=SQL+' WHERE ' + LWhereBase(True,True,False,False) + LWhereComptes(Self) ;
//QR.SQL.Add(SQL) ;
//QR.SQL.Add(LWhereBase(True,True,False,False)) ;
//QR.SQL.Add(LWhereComptes(Self)) ;
SLibres:=GetCritLibres ; if SLibres<>'' then SQL:=SQL+ SLibres ;
SReste:=GetCritReste ; if SReste<>'' then SQL:=SQL+ SReste ;
if Not LETTREPARTIEL.Checked then SQL:=SQL+' AND E_ETATLETTRAGE<>"PL"' ;
if T_NATUREAUXI.Value<>'' then SQL:=SQL+' AND T_NATUREAUXI="'+T_NATUREAUXI.Value+'"' ;
If MonProfilOk(Qui) Then
  BEGIN
  St:=WhereProfilUser(lQ,Qui) ; SQL:=SQL+' AND '+St ;
  END ;
if CAnc.Checked then
 begin
  SQL:=SQL+' AND E_MODESAISIE="-" and MR_NOMBREECHEANCE=1 and E_DEVISE="'+V_PGI.DevisePivot+'" '  ;
  if not (CAvecOD.Checked) then SQL:=SQL+' AND E_NATUREPIECE<>"OD" '  ;
 end ;
SQL:=SQL+'Group by E_GENERAL, E_AUXILIAIRE, E_DEVISE' ;
lQ:=OpenSQL(SQL,true) ;
//ChangeSQL(QR) ; QR.Open ;
if Not lQ.EOF then ChargeListeBase(lQ) else BEGIN Result:=False ; HRappro.Execute(1,'','') ; END ;
Ferme(lQ) ;
END ;

{=========================== DIVERS, UTILITAIRES, AFFICHAGES ============================}
function TFRappAuto.ArretRappro : boolean ;
BEGIN
Result:=False ;
if Not Arret then Exit ; Arret:=False ;
if HRappro.Execute(2,'','')<>mrYes then Exit ;
Result:=True ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/03/2002
Modifié le ... : 14/04/2004
Description .. : LG* en mode pcl, on ne pose pas la question pour le 
Suite ........ : lancement du traitement et l'impression du rapport
Suite ........ : 
Suite ........ : - 12/02/2003 FB 11768 - reaffichage du message de fin de 
Suite ........ : traitement
Suite ........ : -14/04/2004 - LG - lettrage par anc.
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.BValideClick(Sender: TObject);
Var i : integer ;
begin
if OkTraite then Exit ;
OkTraite:=False ;
if ((Not CRef.Checked) and (Not CMontant.Checked) and (Not CSolde.Checked) and (Not CAnc.Checked)) then BEGIN HRappro.Execute(3,'','') ; Exit ; END ;
if ((IFL.RefFC='') or (IFL.RefRC='') or (IFL.RefFF='') or (IFL.RefRF='')) then
   BEGIN
   HRappro.Execute(6,'','') ; Exit ;
   END ;
if CAnc.Checked and ( PGIAsk('Vous allez procéder à un lettrage par ancienneté. Les écritures de règlements seront éclatées en fonction des pièces à lettrer. Voulez-vous continuer ?') = mrNo ) then exit ;
if not ( ctxPCL in V_PGI.PGIContexte) and (HRappro.Execute(0,'','')<>mrYes) then Exit ;
Application.ProcessMessages ;
if Not RechercheComptes then Exit ;
GParam.Enabled:=False ; Pages.Enabled:=False ;
InitMove(TBase.Count,'') ; Arret:=False ; OkTraite:=True ;
for i:=0 to TBase.Count-1 do
    BEGIN
    MoveCur(True) ; if ArretRappro then Break ;
    GeneInd:=i ; if Transactions(LettrerUnCompte,1)<>oeOk then BEGIN MessageAlerte(HRappro.Mess[5]) ; Break ; END ;
    Application.ProcessMessages ;
    END ;
FiniMove ;
GParam.Enabled:=True ; Pages.Enabled:=True ;
{$IFDEF TT}
//VisuPiecesGenere(TBase,EcrGen,6)
{$ELSE}
if ( HRappro.Execute(4,'','')=mrYes ) then
{$IFNDEF GCGC}
   VisuPiecesGenere(TBase,EcrGen,6)
{$ENDIF}
{$ENDIF}
   ;
OkTraite:=False ;
end;

{=============================== RAPPRO REFERENCE ======================================}
function TFRappAuto.BonneDateRef ( LaDate,DateTest : TDateTime ; Client : boolean ) : Boolean ;
BEGIN
if Client then BonneDateRef:=Abs(LaDate-DateTest)<=IFL.TolC
          else BonneDateRef:=Abs(LaDate-DateTest)<=IFL.TolF ;
END ;

function TFRappAuto.MemeRef ( LaRef,RefTest : String ; Client : boolean ) : Boolean ;
Var Strict : boolean ;
BEGIN
if Client then Strict:=IFL.EgalC else Strict:=IFL.EgalF ;
if Strict then MemeRef:=(Laref=RefTest) else MemeRef:=((Pos(LaRef,RefTest)>0) or (Pos(RefTest,LaRef)>0)) ;
END ;

Function TFRappAuto.isISF ( L : TL_Rappro ) : Tisf ;
Var Q : TQuery ;
    Nat : String3 ;
BEGIN
Result:=isfDiv ;
if L.Auxiliaire<>'' then
   BEGIN
   Q:=OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE="'+L.Auxiliaire+'"',True) ;
   if Not Q.EOF then Nat:=Q.Fields[0].AsString else Nat:='' ;
   if ((Nat='CLI') or (Nat='AUD')) then Result:=isfCli else
    if ((Nat='FOU') or (Nat='AUC')) then Result:=isfFou else
     if ((Nat='SAL') or (Nat='DIV')) then BEGIN if CODSALFOU.Checked then Result:=isfFou else Result:=isfCli ; END ;
   Ferme(Q) ;
   END else if L.General<>'' then
   BEGIN
   Q:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+L.General+'"',True) ;
   if Not Q.EOF then Nat:=Q.Fields[0].AsString else Nat:='' ;
   if Nat='TID' then Result:=isfCli else if Nat='TIC' then Result:=isfFou ;
   Ferme(Q) ;
   END ;
END ;

function TFRappAuto.QuelleRef ( L : TL_Rappro ) : String ;
Var Code : String3 ;
    St   : String ;
    isf : Tisf ;
BEGIN
if ((L.Nature='OD') and (CAvecOD.Checked)) then isf:=isISF(L) else isf:=isfDiv ;
if ((L.Nature='FC') or (L.Nature='RC') or (L.Nature='AC') or (L.Nature='OC') or (isf=isfCli)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AC')) then Code:=IFL.RefFC else Code:=IFL.RefRC ;
   if isf=isfCli then
      BEGIN
      if ((L.Debit>0) or (L.Credit<0)) then Code:=IFL.RefFC else Code:=IFL.RefRC ;
      END ;
   END else if ((L.Nature<>'OD') or (isf=isfFou)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AF')) then Code:=IFL.RefFF else Code:=IFL.RefRF ;
   if isf=isfFou then
      BEGIN
      if ((L.Credit>0) or (L.Debit<0)) then Code:=IFL.RefFF else Code:=IFL.RefRF ;
      END ;
   END ;
if Code='LIB' then St:=L.Lib else
 if Code='NUM' then St:=IntToStr(L.Numero) else
  if Code='RFE' then St:=L.RefE else
   if Code='RFI' then St:=L.RefI else
    if Code='RFL' then St:=L.RefL else St:='' ;
QuelleRef:=TripoteSt(St) ;
END ;

function TFRappAuto.QuelleDate ( L : TL_Rappro ; Var ControleDate : boolean ) : TDateTime ;
Var Code : String3 ;
    DD   : TDateTime ;
    isf : Tisf ;
BEGIN
if ((L.Nature='OD') and (CAvecOD.Checked)) then isf:=isISF(L) else isf:=isfDiv ;
DD:=0 ; ControleDate:=True ;
if ((L.Nature='FC') or (L.Nature='RC') or (L.Nature='AC') or (L.Nature='OC')or (isf=isfCli)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AC')) then Code:=IFL.DateFC else Code:=IFL.DateRC ;
   if isf=isfCli then
      BEGIN
      if ((L.Debit>0) or (L.Credit<0)) then Code:=IFL.DateFC else Code:=IFL.DateRC ;
      END ;
   END else if ((L.Nature<>'OD') or (isf=isfFou)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AF')) then Code:=IFL.DateFF else Code:=IFL.DateRF ;
   if isf=isfFou then
      BEGIN
      if ((L.Credit>0) or (L.Debit<0)) then Code:=IFL.DateFF else Code:=IFL.DateRF ;
      END ;
   END ;
if Code='CPT' then DD:=L.DateC else
 if Code='ECH' then DD:=L.DateE else
  if Code='REF' then DD:=L.DateR else ControleDate:=False ;
Result:=DD ;
END ;

procedure TFRappAuto.AnnuleSolutionRef ( NoSol : integer )  ;
Var i : integer ;
    X : TL_Rappro ;
BEGIN
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(Tlett[i]) ; if X.Solution=NoSol then X.Solution:=0 ;
    END ;
END ;

procedure TFRappAuto.SolutionsRef ( NumSol : integer )  ;
Var NoSol,i : integer ;
    X       : TL_Rappro ;
    CodeL,Nat    : String4 ;
    Debit,Credit : double ;
    Total,Coher,Okok,Premier : boolean ;
BEGIN
if NumSol<=0 then Exit ;
for NoSol:=1 to NumSol do
    BEGIN
    CodeL:='' ; Debit:=0 ; Credit:=0 ; Coher:=False ; Premier:=True ;
    for i:=0 to TLett.Count-1 do
        BEGIN
        X:=TL_Rappro(Tlett[i]) ; if X.Solution<>NoSol then Continue ;
        if Premier then
           BEGIN
           Nat:=X.Nature ; Premier:=False ;
           END else
           BEGIN
           if X.Nature<>Nat then Coher:=True ;
           if ((X.Nature='OD') and (CAvecOD.Checked)) then Coher:=True ;
           END ;
        Debit:=Debit+X.DebitCur ; Credit:=Credit+X.CreditCur ;
        if ((CodeL='') and (X.CodeL<>'')) then CodeL:=X.CodeL ;
        END ;
    Total:=Arrondi(Debit-Credit,DEVI.Decimale)=0 ;
    if IFL.LetTotal // FQ13779 SBO 04/08/2004 Prise en Compte du Param Soc "Autoriser le lettrage partiel"
       then Okok:=((CodeL='') and (Total))
       else Okok:=((CodeL='') or (Total)) ;
    Okok:=((Okok) and ((Debit<>0) or (Credit<>0))) ;
    Okok:=((Okok) and (Coher)) ;
    if Okok then FichierLettrage(NoSol,Total,'R') else AnnuleSolutionRef(NoSol) ;
    END ;
END ;

Function TFRappAuto.BonneNature ( Nat1,Nat2 : String ) : boolean ;
BEGIN
if (Nat1='FC') or (Nat1='AC') or (Nat1='RC') or (Nat1='OC') or ((Nat1='OD') and (CAvecOD.Checked))
   then Result:=(Nat2='FC') or (Nat2='AC') or (Nat2='RC') or (Nat2='OC') or ((Nat2='OD') and (CAvecOD.Checked))
   else if (Nat1='FF') or (Nat1='AF') or (Nat1='RF') or (Nat1='OF') or ((Nat1='OD') and (CAvecOD.Checked))
        then Result:=(Nat2='FF') or (Nat2='AF') or (Nat2='RF') or (Nat2='OF') or ((Nat2='OD') and (CAvecOD.Checked))
             else Result:=False ;
END ;

procedure TFRappAuto.LettrerReference ;
Var i,j,NumSol : integer ;
    L,X        : TL_Rappro ;
    LaRef,RefTest : String ;
    LaDate,DateTest : TDateTime ;
    Client,Okok,ControleDate,b,Premier : Boolean ;
BEGIN
NumSol:=0 ;
for i:=0 to TLett.Count-1 do
    BEGIN
    L:=TL_Rappro(TLett[i]) ; if L.Solution>0 then Continue ; Premier:=True ;
    Client:=L.Client ;
    if Client then Okok:=(L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC') or ((L.Nature='OD') and (CAvecOD.Checked))
              else Okok:=(L.Nature='FF') or (L.Nature='AF') or (L.Nature='RF') or (L.Nature='OF') or ((L.Nature='OD') and (CAvecOD.Checked)) ;
    if Okok then
       BEGIN
       LaRef:=QuelleRef(L) ; if LaRef='' then Continue ;
       LaDate:=QuelleDate(L,ControleDate) ;
       for j:=0 to TLett.Count-1 do if j<>i then
           BEGIN
           X:=TL_Rappro(TLett[j]) ; if X.Solution>0 then Continue ;
           if Not BonneNature(L.Nature,X.Nature) then Continue ;
           RefTest:=QuelleRef(X) ; if RefTest='' then Continue ;
           if ControleDate then
              BEGIN
              DateTest:=QuelleDate(X,b) ;
              if Not BonneDateRef(LaDate,DateTest,L.Client) then Continue ;
              END ;
           if MemeRef(LaRef,RefTest,L.Client) then
              BEGIN
              if Premier then BEGIN Inc(NumSol) ; Premier:=False ; END ;
              L.Solution:=NumSol ; X.Solution:=NumSol ;
              END ;
           END ;
       END ;
    END ;
SolutionsRef(NumSol) ;
END ;

{================================ LETTRAGE EPUISEMENT ===================================}
{=========================== LETTRAGE MONTANT COMBINATOIRE ===========================}
function TFRappAuto.ConstruitListes ( icur : integer ; Debit : boolean ; Var LM : T_D ; Var LP,LG : T_I ) : integer ;
Var j,indice : integer ;
    X : TL_Rappro ;
BEGIN
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ; FillChar(LG,Sizeof(LG),#0) ; Indice:=0 ;
for j:=0 to TLett.Count-1 do if j<>icur then
    BEGIN
    X:=TL_Rappro(TLett[j]) ;
    if X.Solution=0 then
       BEGIN
       if Debit then
          BEGIN
          if X.DebitCur<>0 then LM[Indice]:=-X.DebitCur else LM[Indice]:=X.CreditCur ;
          END else
          BEGIN
          if X.CreditCur<>0 then LM[Indice]:=-X.CreditCur else LM[Indice]:=X.DebitCur ;
          END ;
       LP[Indice]:=0 ; LG[Indice]:=j ; Inc(Indice) ;
       if Indice>=MaxDroite then Break ;
       END ;
    END ;
Result:=Indice ;
END ;

procedure TFRappAuto.LettrerMontant ;
Var i,Res,NumSol,NbD,k : integer ;
    L     : TL_Rappro ;
    LM    : T_D ;
    LP,LG : T_I ;
    Debit : boolean ;
    Solde : double ;
    Infos : REC_AUTO ;
    TT1   : TDateTime ;
    NbS   : Longint ;
    HH,MM,SS,CC : Word ;
BEGIN
TT1:=Time ;
NumSol:=GeneNumSol ;
for i:=0 to TLett.Count-1 do
    BEGIN
    L:=TL_Rappro(TLett[i]) ; if L.Solution<>0 then Continue ;
    if (L.Facture) or ((L.Nature='OD') and (CAvecOD.Checked)) then
       BEGIN
       Debit:=(L.DebitCur<>0) ; if Debit then Solde:=L.DebitCur else Solde:=L.CreditCur ;
       NbD:=ConstruitListes(i,Debit,LM,LP,LG) ;
       Infos.Nival:=MAXPROF.Value-1 ; Infos.NbD:=NbD ; Infos.Decim:=DEVI.Decimale ;
       Infos.Temps:=MAXDUREE.Value ; Infos.Unique:=False ;
       Res:=LettrageAuto(Solde,LM,LP,Infos) ;
       if Res=1 then
          BEGIN
          Inc(NumSol) ; L.Solution:=NumSol ;
          for k:=0 to NbD do if LP[k]<>0 then TL_Rappro(TLett[LG[k]]).Solution:=NumSol ;
          FichierLettrage(NumSol,True,'M') ;
          END ;
       END ;
    DecodeTime(Time-TT1,HH,MM,SS,CC) ;
    NbS:=3600*HH+60*MM+SS ; if NbS>MaxDuree.Value then Break ;
    END ;
END ;

{================================ LETTRAGE SOLDE NUL =================================}
procedure TFRappAuto.LettrerSolde ;
Var i,k,NumSol,Depart : integer ;
    X : TL_Rappro ;
    Solde : double ;
    PasDeSolution : boolean ;
BEGIN
NumSol:=GeneNumSol ; Depart:=0 ;
Repeat
 Solde:=0 ; PasDeSolution:=True ;
 for i:=Depart to TLett.Count-1 do
     BEGIN
     X:=TL_Rappro(TLett[i]) ; if X.Solution<>0 then Continue ;
     Solde:=Solde+X.DebitCur-X.CreditCur ;
     if Arrondi(Solde,DEVI.Decimale)=0 then
        BEGIN
        Inc(NumSol) ; PasDeSolution:=False ;
        for k:=Depart to i do if TL_Rappro(TLett[k]).Solution=0 then TL_Rappro(TLett[k]).Solution:=NumSol ;
        FichierLettrage(NumSol,True,'S') ;
        Depart:=i+1 ; Break ;
        END ;
     END ;
Until ((Depart>=TLett.Count) or (PasDeSolution)) ;
END ;

{================================ LETTRAGE FICHIER ===================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : - LG - 12/04/2007 - affectation du e_paquetrevision
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.FinirLettrage ( NoSol : integer ; Total : boolean ; DMin,DMax : TDateTime ; CodeL : String ; ModeRappro : Char ) ;
Var i : integer ;
    X : TL_Rappro ;
    B : T_RAPPAUTO ;
    CD,CP,TotCDebit,TotCCredit : Double ;
    SQL,EL,LDEV,STE : String ;
BEGIN
if Total then BEGIN EL:='TL' ; CodeL:=uppercase(CodeL) ; END else BEGIN EL:='PL' ; CodeL:=LOWERCASE(CodeL) ; END ;
{Analyse lettrage monnaie de tenue}
TotCDebit:=0 ; TotCCredit:=0 ;
if DEVI.Code<>V_PGI.DevisePivot then
   BEGIN
   for i:=0 to TLett.Count-1 do
       BEGIN
       X:=TL_Rappro(TLett[i]) ; if X.Solution<>NoSol then Continue ;
       CD:=X.CouvCur ;
       CP:=DeviseToEuro(CD,X.TauxDev,DEVI.Quotite) ;
       if X.DebitCur<>0 then TotCDebit:=TotCDebit+CP else TotCCredit:=TotCCredit+CP ;
       END ;
   END ;
if Arrondi(TotCDebit-TotCCredit,V_PGI.OkDecV)=0 then StE:='---0A'+ModeRappro+'0000'
                                             else StE:='--X0A'+ModeRappro+'0000' ;
{Lettrage}
if DEVI.Code<>V_PGI.DevisePivot then LDEV:='X' else LDEV:='-' ;
B:=T_RAPPAUTO(TBase[CurBase]) ;
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(TLett[i]) ; if X.Solution<>NoSol then Continue ;
    CD:=X.CouvCur ;
    if Devi.Code<>V_PGI.DevisePivot then
     CP:=DeviseToEuro(CD,X.TauxDev,DEVI.Quotite) 
      else
       CP:=CD ;

    SQL:='UPDATE ECRITURE SET E_PAQUETREVISION=1,E_LETTRAGE="'+CodeL+'", E_ETATLETTRAGE="'+EL+'", E_DATEPAQUETMIN="'+USDATETIME(DMin)+'", '
        +'E_DATEPAQUETMAX="'+USDATETIME(DMax)+'", E_COUVERTURE='+StrFPoint(CP)+', E_COUVERTUREDEV='+StrFPoint(CD)+', '
        +'E_LETTRAGEDEV="'+LDEV+'" ,E_ETAT="'+StE+'", E_REFLETTRAGE="'+E_REFLETTRAGE.Text+'", E_TRESOSYNCHRO = "LET" ' ;{JP 27/04/04 : pour l'échéancier de la Tréso}
    {#TVAENC}
    if VH^.OuiTvaEnc then
       BEGIN
       if RazTvaEnc(X) then SQL:=SQL+', E_ECHEENC1=0, E_ECHEENC2=0, E_ECHEENC3=0, E_ECHEENC4=0, E_ECHEDEBIT=0 ' ;
       END ;
    SQL:=SQL+'Where E_GENERAL="'+X.General+'" AND E_AUXILIAIRE="'+X.Auxiliaire+'" AND E_EXERCICE="'+X.Exo+'" '
            +'AND E_JOURNAL="'+X.Jal+'" AND E_DATECOMPTABLE="'+USDATETIME(X.DateC)+'" AND E_NUMEROPIECE='+IntToStr(X.Numero)+' '
            +'AND E_NUMLIGNE='+IntToStr(X.NumLigne)+' AND E_NUMECHE='+IntToStr(X.NumEche) ;
    if ExecuteSQL(SQL)<>1 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
    if X.DebitCur<>0 then BEGIN Inc(B.NbD) ; B.TotalD:=B.TotalD+X.DebitCur ; END
                     else BEGIN Inc(B.NbC) ; B.TotalC:=B.TotalC+X.CreditCur ; END ;
    B.CodeZ:=CodeL ; if B.CodeA='' then B.CodeA:=CodeL ;
    END ;
END ;

procedure TFRappAuto.DelettrePartiel ( X : TL_Rappro ) ;
Var SQL : String ;
BEGIN
SQL:='UPDATE ECRITURE SET E_LETTRAGE="", E_ETATLETTRAGE="AL", E_COUVERTURE=0, E_COUVERTUREDEV=0, ' ;
SQL:=SQL+'E_LETTRAGEDEV="-", E_REFLETTRAGE="", E_TRESOSYNCHRO = "LET" ' ;{JP 27/04/04 : pour l'échéancier de la Tréso}
SQL:=SQL+'Where E_AUXILIAIRE="'+X.Auxiliaire+'" AND E_GENERAL="'+X.General+'" ' ;
SQL:=SQL+'AND E_ETATLETTRAGE="PL" AND E_LETTRAGE="'+X.CodeL+'"' ;
ExecuteSQL(SQL) ;
END ;

procedure TFRappAuto.FichierLettrage ( NoSol : integer ; Total : boolean ; ModeRappro : Char ) ;
Var i : integer ;
    X : TL_Rappro ;
    Premier   : boolean ;
    CodeL     : String ;
    LD,LC     : TList ;
    DMin,DMax : TDateTime ;
BEGIN
Premier:=True ; LD:=TList.Create ; LC:=TList.Create ; if NoSol>GeneNumSol then GeneNumSol:=NoSol ;
DMin:=0 ; DMax:=0 ;
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(TLett[i]) ; if X.Solution<>NoSol then Continue ;
    if ((Total) and (X.CodeL<>'')) then DelettrePartiel(X) ;
    if Premier then
       BEGIN
       CodeL:=GetSetCodeLettre(X.General,X.Auxiliaire) ; Premier:=False ;
       DMin:=X.DateC ; DMax:=X.DateC ;
       END else
       BEGIN
       if DMin>X.DateC then DMin:=X.DateC ;
       if DMax<X.DateC then DMax:=X.DateC ;
       END ;
    if Total then
       BEGIN
       if X.DebitCur<>0 then X.CouvCur:=X.DebitCur else X.CouvCur:=X.CreditCur ;
       END else
       BEGIN
       if X.DebitCur>0 then LD.Add(TLett[i]) else LC.Add(TLett[i]) ;
       END ;
    END ;
if Not Total then RefEpuise(LD,LC) ;
FinirLettrage(NoSol,Total,DMin,DMax,CodeL,ModeRappro) ;
LD.Clear ; LD.Free ; LC.Clear ; LC.Free ;
END ;

procedure TFRappAuto.RecupInfosDevise ;
Var ii : integer ;
    St,St1 : String ;
BEGIN
ii:=CDEV.Items.IndexOf(DEVI.Code) ;
if ii>=0 then
   BEGIN
   St:=CDEV.Values[ii] ;
   St1:=ReadTokenSt(St) ; DEVI.Decimale:=StrToInt(St1) ;
   St1:=ReadTokenSt(St) ; DEVI.Quotite:=StrToFloat(St1) ;
   END else if DEVI.Code=V_PGI.DevisePivot then
   BEGIN
   DEVI.Decimale:=V_PGI.OkDecV ; DEVI.Quotite:=1 ;
   END ;
END ;

procedure TFRappAuto.LettrerUnCompte ;
Var T : T_RAPPAUTO ;
    Ind : integer ;
BEGIN
try
Ind:=GeneInd ; T:=T_RAPPAUTO(TBase[ind]) ; CurBase:=Ind ;
if ((T.Devise<>DEVI.Code) or (Ind=0)) then BEGIN DEVI.Code:=T.Devise ; RecupInfosDevise ; END ;
ChargeEcrCompte(T) ; GeneNumSol:=0 ;
if CRef.Checked then LettrerReference ;
if CMontant.Checked then LettrerMontant ;
if CSolde.Checked then LettrerSolde ;
if CAnc.Checked then LettrerAnc ;
except
 On E : Exception do
  begin
   MessageAlerte('Erreur' + #10#13 + E.Message ) ;
   raise ;
  end ;
end ;
END ;

procedure TFRappAuto.CMontantClick(Sender: TObject);
Var Ena : boolean ;
begin
Ena:=CMontant.Checked ;
MaxProf.Enabled:=Ena ; H_MaxProf.Enabled:=Ena ;
MaxDuree.Enabled:=Ena ; H_MaxDuree.Enabled:=Ena ;
end;

procedure TFRappAuto.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 02/11/2004
Modifié le ... :   /  /
Description .. : - LG - 02/11/2004 - FB 14434 - le lettrage partiel est 
Suite ........ : forcement active pour cette version
Mots clefs ... : 
*****************************************************************}
procedure TFRappAuto.CRefClick(Sender: TObject);
begin
CODSALFOU.Enabled:=(CRef.Checked) and (CAvecOD.Checked) ;
LETTREPARTIEL.checked:=CAnc.Checked ; 
LETTREPARTIEL.enabled:= not CAnc.Checked ;
end;

procedure TFRappAuto.CAvecODClick(Sender: TObject);
begin
CODSALFOU.Enabled:=(CRef.Checked) and (CAvecOD.Checked) ;
end;

procedure TFRappAuto.E_EXERCICEChange(Sender: TObject);
begin
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFRappAuto.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFRappAuto.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Key=VK_F10 then BEGIN Key:=0 ; BValideClick(Nil) ; END ; 
end;

procedure TFRappAuto.LettrerAnc ;
var
 RL : RLETTR ;
 L : TL_Rappro ;
begin
 DEVI.Taux:=GetTaux(DEVI.Code,DEVI.DateTaux,V_PGI.DateEntree) ;
 L := TL_Rappro(Tlett[0]) ;
 CLettrageParAnc( ((L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) , RL, TLett,DEVI , T_RAPPAUTO(TBase[CurBase]) ) ;
 RegulLettrage (true,false,true) ;
end ;

end.
