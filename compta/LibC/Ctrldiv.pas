unit CtrlDiv;

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
  StdCtrls,
  Buttons,
  ExtCtrls,
  Ent1,
  HEnt1,
  Hctrls,
  hmsgbox,
  HSysMenu,
  SoldeCpt,
  HPanel,
  UiUtil,
  HTB97,
{$IFDEF NETEXPERT}
  Paramsoc,
  UtilTrans,
{$ENDIF}
  RepEntete
  ;

procedure ReparationFic ;



type
  TFCtrlDiv = class(TForm)
    Panel1: TPanel;
    FControleMvt: TGroupBox;
    FDatePaquet: TCheckBox;
    MsgRien: THMsgBox;
    FcontroleCpt: TGroupBox;
    FGen: TCheckBox;
    FAux: TCheckBox;
    FSec: TCheckBox;
    FJal: TCheckBox;
    FSolde: TCheckBox;
    FANouveau: TCheckBox;
    Mes: THMsgBox;
    HMTrad: THSystemMenu;
    FNumPL: TCheckBox;
    MsgBar: THMsgBox;
    FLettrage: TCheckBox;
    FContreparties: TCheckBox;
    FReparMvt: TCheckBox;
    TTravail: TLabel;
    TTravailLeq: TLabel;
    FDevMnt: TCheckBox;
    FPeriode: TCheckBox;
    FLettRef: TCheckBox;
    FTvaEnc: TCheckBox;
    FCodeAccept: TCheckBox;
    FSoldeP: TCheckBox;
    FEntetePiece: TCheckBox;
    Dock972: TDock97;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    ANNREVISION: TCheckBox;
    ckAuxiAna: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées }
    Piece, Solde, DatPaq,Lettre,ReparMvt,
    ANouveau,Contrepartie,Cpt,ReparDev : Boolean ;
    bPeriode,LettreCode,TvaEnc,CodeAccept,Entete: Boolean ;
  // GP le 23/05/2008 : réactivé pour cause de pointage :
    SoldeP : Boolean;
    {SoldeP : Boolean; JP 17/01/08 : FQ 22222 : les totaux pointés ne servent plus à rien}
    Procedure LanceControl ;
    Procedure AvecQui ;
    Function  Selection : Boolean ;
    Procedure MajCaption( M : byte ) ;
  public
    { Déclarations publiques }
    {JP 09/11/07 : FQ 21813 : réparation des auxiliaire sur l'analytique}
    procedure ReparAuxiAna;
  end;


implementation

{$R *.DFM}

uses
  uTob, HStatus, {JP 09/11/07 : ReparAuxiAna}
  Fe_Main, VerCpte, VerPiece, VerCpta, CPTESAV, CPTEUTIL,
  VerDPaq,
  ReCalcTva,
  TofMajPerSem,
  uTofCPOutilLettrage,
  // GP le 23/05/2008 : réactivé pour cause de pointage :
  CalcSoldeP,
  CPAVERTNEWRAPPRO_TOF,
  {CalcSoldeP, JP 17/01/08 : FQ 22222 : les totaux pointés ne servent plus à rien}
  VerLettr,VerContr, ReparMvt, RepDevEur, LETTREF, ReinitCA,
  UtilPgi;

procedure ReparationFic ;
var CDiv : TFCtrlDiv ;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(True,'',TRUE) then Exit ;
CDiv:=TFCtrlDiv.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     CDiv.ShowModal ;
    finally
     CDiv.Free ;
     _DeblocageMonoPoste(True,'',TRUE) ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(CDiv,PP) ;
   CDiv.Show ;
   END ;
END ;

procedure TFCtrlDiv.FormShow(Sender: TObject);
{$IFDEF NETEXPERT}
var
  LienInterGamme : string;
{$ENDIF}
begin

PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TTravail.Caption:='' ; TTravailLeq.Caption:='' ;

// Rony 23/05/97 Anou fonction non dispo encore
FANouveau.Checked:=False ; FANouveau.Enabled:=False ; FDevMnt.Enabled:=FALSE ;
FANouveau.Visible:=V_PGI.SAV ;
FNumPL.Visible:=V_PGI.SAV ;

// GCO - 04/08/2006 - Branchement Outil de correction du Lettrage
FLettrage.Enabled :=True;
FLettrage.Visible :=True;
// FIN GCO

{$IFDEF NETEXPERT}
LienInterGamme := GetParamsoc ('SO_CPLIENGAMME');
ANNREVISION.visible := (LienInterGamme<>'') and (LienInterGamme<>'SI') and (LienInterGamme<>'AUC') ;
{$ENDIF}
If Not OkNewPointage Then
  BEGIN
  FSoldeP.Visible:=TRUE ;
  FSoldeP.Enabled:=TRUE ;
  END Else
  BEGIN
  FSoldeP.Visible:=FALSE ;
  FSoldeP.Checked:=FALSE ;
  END ;
end;

procedure TFCtrlDiv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True,'',TRUE) ;
   Action:=caFree ;
   END ;
end;

procedure TFCtrlDiv.BValiderClick(Sender: TObject);
begin
if Not Selection then begin MsgRien.Execute(0,'','') ; Exit ; end ;
if MsgRien.Execute(2,'','')<>mryes then exit ;
TTravail.Caption:=Mes.Mess[4] ;
AvecQui ;
LanceControl ;
TTravail.Caption:='' ;
TTravailLeq.Caption:='' ;
end;

Procedure TFCtrlDiv.AvecQui ;
BEGIN
Cpt:=(FGen.Checked or FAux.Checked or FSec.Checked or FJal.Checked or ANNREVISION.Checked) ;
Piece:=FNumPL.Checked ;
Solde:=FSolde.Checked ;
ANouveau:=FANouveau.Checked ;
DatPaq:=FDatePaquet.Checked ;
Lettre:=FLettrage.Checked ;
Contrepartie:=FContreparties.Checked ;
ReparMvt:=FReparMvt.Checked ;
ReparDev:=FDevMnt.Checked ;
bPeriode:=FPeriode.Checked ;
LettreCode:=FLettRef.Checked ;
TvaEnc:=FTvaEnc.Checked ;
CodeAccept:=FCodeAccept.Checked ;
  // GP le 23/05/2008 : réactivé pour cause de pointage :
SoldeP:=FSoldeP.Checked ;
{SoldeP:=FSoldeP.Checked ; JP 17/01/08 : FQ 22222 : les totaux pointés ne servent plus à rien}
Entete:=FEntetePiece.Checked ;
END ;

Procedure TFCtrlDiv.LanceControl ;
Var  NbNoPossible : integer ;
BEGIN
EnableControls(Self,False) ;
if Cpt then
   BEGIN
   if FGen.Checked then
      BEGIN
      NbNoPossible:=0 ; MajCaption(1) ;
      VerCompteMAJ(1,NbNoPossible, False) ;
      if NbNoPossible=1 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[0]) ;
      if NbNoPossible>1 then MsgRien.Execute(4,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[0]) ;
      END ;
   if FAux.Checked then
      BEGIN
      NbNoPossible:=0 ; MajCaption(2) ;
      VerCompteMAJ(2,NbNoPossible, False) ;
      if NbNoPossible=1 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[1]) ;
      if NbNoPossible>1 then MsgRien.Execute(4,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[1]) ;
      END ;
   if FSec.Checked then
      BEGIN
      NbNoPossible:=0 ; MajCaption(3) ;
      VerCompteMAJ(3,NbNoPossible, False) ;
      if NbNoPossible=1 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[2]) ;
      if NbNoPossible>1 then MsgRien.Execute(4,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[2]) ;
      END ;
   if FJal.Checked then
      BEGIN
      NbNoPossible:=0 ; MajCaption(4) ;
      VerCompteMAJ(4,NbNoPossible, False) ;
      if NbNoPossible=1 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[3]) ;
      if NbNoPossible>1 then MsgRien.Execute(4,IntToStr(NbNoPossible)+' ',' '+Mes.Mess[3]) ;
      END ;
{$IFDEF NETEXPERT}
   if ANNREVISION.Checked then  // ajout me 03/08/2005
        UpdateComLettrage(8);
{$ENDIF}
   END ;
if ANouveau then begin MajCaption(6) ; VerifAno(False) ; end ;
if DatPaq then begin MajCaption(7) ; DatePaquet ; end ;
//if DatPaq then begin MajCaption(7) ; SAVDATEMINDATEMAXPAQUET(Tous) ; end ;
//if Lettre then begin MajCaption(8) ; RepareLeLettrage ; end ;
if Lettre then begin MajCaption(8) ; CPLanceFiche_CPOutilLettrage; end ;
if Piece then begin MajCaption(9) ; VerifPiece ; end ;
If Contrepartie Then begin MajCaption(11) ; VerifContreparties ; end ;
If ReparMvt Then RepareMvt ;
if Solde then begin MajCaption(5) ; MajTotTousComptes(False,'') ; end ;
If ReparDev Then ReparationMontantEuro ;
if bPeriode then
  begin
  MajCaption(13) ;
  AGLLanceFiche('CP', 'MAJPERIOSEMAIN', '', '', '') ;
  end ;
//if DatPaq then SAVDATEMINDATEMAXPAQUET ;
If LettreCode Then begin LettrageParCode ; end ;
If TvaEnc Then CalcLaTva ;
If CodeAccept Then ReinitCodeAccept ;

// GP le 23/05/2008 : réactivé pour cause de pointage :
If SoldeP Then CalcTotP ;
{If SoldeP Then CalcTotP ; JP 17/01/08 : FQ 22222 : les totaux pointés ne servent plus à rien}
if Entete then ReparationEntete;
{JP 09/11/07 : FQ 21813 : réparation des auxiliaire sur l'analytique}
if ckAuxiAna.Checked then ReparAuxiAna;

EnableControls(Self,True) ;
Screen.Cursor:=SyncrDefault ;
MsgRien.Execute(1,'','') ;
END ;

Function TFCtrlDiv.Selection : Boolean ;
BEGIN
Result:=(FGen.Checked or FAux.Checked or FSec.Checked or FJal.Checked
         or FNumPL.Checked or FSolde.Checked or FDatePaquet.Checked or FLettrage.Checked
         or FANouveau.Checked or FContreparties.Checked Or FReparMvt.Checked
         or FDevMnt.Checked or FPeriode.Checked Or FLettRef.Checked Or FTvaEnc.Checked
         or FCodeAccept.Checked or FEntetePiece.Checked or ANNREVISION.checked
// GP le 23/05/2008 : réactivé pour cause de pointage :
         or FSoldeP.Checked
         {or FSoldeP.Checked JP 17/01/08 : FQ 22222 : les totaux pointés ne servent plus à rien}
         or ckAuxiAna.Checked);{JP 09/11/07 : FQ 21813}
END ;

Procedure TFCtrlDiv.MajCaption( M : byte ) ;
BEGIN
TTravail.Caption:=MsgBar.Mess[0]+' ' ;
TTravailLeq.Caption:=MsgBar.Mess[M] ;
Application.ProcessMessages ;
END ;

procedure TFCtrlDiv.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCtrlDiv.BFermeClick(Sender: TObject);
begin
  //SG6 26/01/05
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

{JP 09/11/07 : FQ 21813 : réparation des auxiliaire sur l'analytique
{---------------------------------------------------------------------------------------}
procedure TFCtrlDiv.ReparAuxiAna;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  n : Integer;
  s : string;
begin
  T := TOB.Create('MOMO', nil, -1);
  try
    T.LoadDetailFromSQL('SELECT EX_EXERCICE FROM EXERCICE');
    InitMove(T.Detail.Count, TraduireMemoire('Réparation des auxiliaires'));
    for n := 0 to T.Detail.Count - 1 do begin
      MoveCur(False);
      s := 'UPDATE ANALYTIQ SET Y_AUXILIAIRE = (SELECT MAX(E_AUXILIAIRE) FROM ECRITURE ';
      s := s + 'WHERE E_JOURNAL = Y_JOURNAL AND E_EXERCICE = Y_EXERCICE AND E_DATECOMPTABLE = Y_DATECOMPTABLE ';
      s := s + 'AND E_NUMEROPIECE = Y_NUMEROPIECE AND E_NUMLIGNE = Y_NUMLIGNE AND E_QUALIFPIECE = Y_QUALIFPIECE) ';
      s := s + 'WHERE Y_EXERCICE = "' + T.Detail[n].GetString('EX_EXERCICE') + '" AND ';
      s := s + '(Y_TYPEANALYTIQUE <> "X" OR Y_TYPEANALYTIQUE IS NULL)';
      ExecuteSQL(s);
      Application.ProcessMessages
    end;

  finally
    FreeAndNil(T);
    FiniMove;
  end;
end;

end.
