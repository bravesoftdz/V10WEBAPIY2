unit QRLegal;

interface

uses
Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
StdCtrls, Buttons, Hctrls, ComCtrls, HQuickRP, ExtCtrls, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
hmsgbox, Ent1, HQry, HEnt1, Mask, Hcompte, CpteUtil, ParamDat, QRRupt,
SaisUtil, UtilEdt, EdtLegal, Filtre, Menus, CritEdt, QR, HSysMenu, HTB97, HPanel, UiUtil ;

Type TDou=(neRien, neAudit) ;

procedure LanceEdtLegal ;
Function  LanceEdtLegClo : TTraitement ;

type
  TEdLegal = class(TForm)
    Pages: TPageControl;
    Standards: TTabSheet;
    GlvAux: TCheckBox;
    BalAux: TCheckBox;
    JalDiv: TCheckBox;
    BalGen: TCheckBox;
    GlvGen: TCheckBox;
    HLabel4: THLabel;
    HLabel3: THLabel;
    TFDateCpta2: TLabel;
    FExercice: THValComboBox;
    Apercu: TCheckBox;
    MsgBleme: THMsgBox;
    FDate1: THValComboBox;
    FDate2: THValComboBox;
    FDateCompta1: TMaskEdit;
    FDateCompta2: TMaskEdit;
    FCouleur: TCheckBox;
    FDateCpta1: THValComboBox;
    FDateCpta2: THValComboBox;
    HMTrad: THSystemMenu;
    PourClotureOk: TCheckBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    Panelfiltre: TToolWindow97;
    HPB: TToolWindow97;
    BFiltre: TToolbarButton97;
    BAide: TToolbarButton97;
    BFerme: TToolbarButton97;
    BValider: TToolbarButton97;
    BStop: TToolbarButton97;
    FFiltres: THComboBox;
    procedure FExerciceChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FDateCpta1Change(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    Crit          : TCritEdt ;
    Continuer, OrigineColor,
    LoadFil    : Boolean ;
    Dou        : TDou ;
    Procedure RempliComboExo ;
    Function  PeriodesValidees : Boolean ;
    procedure LanceEdit ;
    function  CritEdtOk : Boolean ;
    procedure GenPerExo ;
    procedure ChangeDATECOMPTA ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses QRGLAux, QRGLGen, QRBalAux, QRBalGen, QRJDivis, CpteSav ;

procedure LanceEdtLegal ;
var QR : TEdLegal ;
    PP : THPanel ;
BEGIN
SourisSablier ;
PP:=FindInsidePanel ;
QR:=TEdLegal.Create(Application) ;
QR.Dou:=neRien ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;

Function  LanceEdtLegClo : TTraitement ;
var QR : TEdLegal ;
BEGIN
SourisSablier ; Result:=cOk ;
QR:=TEdLegal.Create(Application) ;
try
 QR.Dou:=neAudit ;
 QR.ShowModal ;
  Case QR.PourClotureOk.State of
    cbGrayed    : Result:=cPasFait ;
    cbChecked   : Result:=cOk ;
    cbUnChecked : Result:=cPasOk ;
    End ;
 finally
 QR.Free ;
 end ;
SourisNormale ;
END ;

procedure TEdLegal.FormShow(Sender: TObject);
begin

// GC - 14/01/2002
if CtxPcl in V_Pgi.PgiContexte then
  JalDiv.Caption := '&Journal des écritures'
else
  JalDiv.Caption := '&Jal divisionnaire';
// GC - FIN

PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
RempliComboExo; OrigineColor:=V_PGI.QRCouleur ;
FCouleur.State:=cbChecked ; PourClotureOk.State:=cbGrayed ;
If Dou=neAudit then FExercice.Value:=VH^.Encours.Code else FExercice.Value:=VH^.Entree.Code ;
Continuer:=True ;
if Dou=neAudit then
   BEGIN
   PanelFiltre.Enabled:=False ;
   END Else
   BEGIN
   LoadFil:=True ;
   ChargeFiltre('EDTLEGAL',FFiltres,Pages) ;
   LoadFil:=False ;
   END ;
SourisNormale ;
end;

FUNCTION DonneAnnee(LaDate : TDateTime) : Integer ;
Var a,m,j : Word ;
BEGIN
DecodeDate(LaDate,a,m,j) ;
Result:=a ;
END ;

Procedure TEdLegal.RempliComboExo ;
Var Q : TQuery ;
    Sql : String ;
BEGIN
If Dou=neAudit then
   BEGIN
   Sql:='Select EX_EXERCICE,EX_LIBELLE from EXERCICE'+
        ' Where EX_EXERCICE="'+VH^.Encours.Code+'" '  ;
   END Else
   BEGIN
   Sql:='Select EX_EXERCICE,EX_LIBELLE from EXERCICE'+
        ' Where EX_EXERCICE="'+VH^.EnCours.Code+'" Or EX_EXERCICE="'+VH^.Suivant.Code+'"'  ;
   END ;
Q:=OpenSql(Sql,True) ;
FExercice.Values.Clear ; FExercice.Items.Clear ;
While Not Q.EOF do
    BEGIN
    FExercice.Values.Add(Q.Fields[0].AsString) ;
    FExercice.Items.Add(Q.Fields[1].AsString) ;
    Q.Next ;
    END ;
Ferme(Q) ;
END ;

procedure TEdLegal.GenPerExo ;
{ Création des périodes comptables pour un exercice donné }
var DateExo            : TExoDate ;
    i                  : integer ;
    Annee,pMois,NbMois : Word ;
    DD : TdateTime ;
    D1, D2 : String ;
BEGIN
NbMois:=0 ;
QuelDateDeExo(FExercice.Value,DateExo) ;
NOMBREPEREXO(DateExo,pMois,Annee,NbMois) ;
If Not LoadFil then
   BEGIN
   FDateCpta1.Items.Clear ; FDateCpta1.Values.Clear ; FDateCpta2.Items.clear ; FDateCpta2.Values.Clear ;
   FDate1.Items.Clear ; FDate1.Values.Clear ; FDate2.Items.Clear ; FDate2.Values.Clear ;
   for i:=0 to NbMois-1 do
       BEGIN
       DD:=PlusMois(DateExo.Deb,i) ;
       D1:=FormatDateTime('mmmm yyyy',DD) ;
       D2:=FormatDateTime('mmmm yyyy',DD) ;
       FDateCpta1.Items.Add(FirstMajuscule(D1));
       FDateCpta2.Items.Add(FirstMajuscule(D2));
       FDateCpta1.Values.Add(IntTostr(i)) ; // Rony 14/05/97 -- A cause du filtre
       FDateCpta2.Values.Add(IntTostr(i)) ; //
       FDate1.Items.Add(DateToStr(DebutdeMois(dd))) ;
       FDate2.Items.Add(DateToStr(FindeMois(dd))) ;
       END ;
  FDateCpta1.ItemIndex:=0 ; FDateCpta2.ItemIndex:=NbMois-1 ;
  END;
FDate1.ItemIndex:=FDateCpta1.ItemIndex; FDate2.ItemIndex:=FDateCpta2.ItemIndex ;
ChangeDateCompta ;
END ;

procedure TEdLegal.ChangeDATECOMPTA ;
BEGIN
FDateCompta1.Text:=FDate1.Items[FDate1.ItemIndex] ;
FDateCompta2.Text:=FDate2.Items[FDate2.ItemIndex] ;
END ;

procedure TEdLegal.FExerciceChange(Sender: TObject);
begin
GenPerExo ;
end;

procedure TEdLegal.BCreerFiltreClick(Sender: TObject);
begin
NewFiltre('EDTLEGAL',FFiltres,Pages) ;
end;

procedure TEdLegal.BDelFiltreClick(Sender: TObject);
begin
DeleteFiltre('EDTLEGAL',FFiltres) ;
end;

procedure TEdLegal.BRenFiltreClick(Sender: TObject);
begin
RenameFiltre('EDTLEGAL',FFiltres) ;
end;

procedure TEdLegal.BSaveFiltreClick(Sender: TObject);
begin
SaveFiltre('EDTLEGAL',FFiltres,Pages) ;
end;

procedure TEdLegal.BNouvRechClick(Sender: TObject);
begin
VideFiltre(FFiltres,Pages) ;
end;

procedure TEdLegal.FFiltresChange(Sender: TObject);
begin
LoadFil:=True ;
LoadFiltre('EDTLEGAL',FFiltres,Pages) ;
LoadFil:=False ;
end;

procedure TEdLegal.BStopClick(Sender: TObject);
begin
Continuer:=False ;
//BStop.Enabled:=False ;
end;

procedure TEdLegal.LanceEdit;
begin
OnImpr:=Not Apercu.Checked ;
V_PGI.QRCouleur:=FCouleur.Checked ;
Crit.SoldeProg:=0 ; Crit.Qualifpiece:='NOR' ;
if (GlvAux.Checked) and (Continuer) then
   BEGIN
   Crit.NatureEtat:=neGL ;
   InitCritEdt(Crit) ;
   Crit.GL.TypCpt:=3 ; Crit.GL.QuelAN:=AvecAN ;
   Crit.ModeRevision:=cbUnChecked ; Crit.GL.TotEche:=True ;
   GLAuxiliaireLegal(Crit) ;
   END ;
if (GlvGen.Checked) and (Continuer) then
   BEGIN
   Crit.NatureEtat:=neGL ;
   InitCritEdt(Crit) ;
   Crit.GL.TypCpt:=3 ; Crit.GL.QuelAN:=AvecAN ;
   Crit.ModeRevision:=cbUnChecked ;
   Crit.GL.ForceNonCentralisable:=False ; Crit.GL.SansDetailCentralisation:=False ;
   GLGeneralLegal(Crit) ;
   END ;
if (BalAux.Checked) and (Continuer) then
   BEGIN
   Crit.NatureEtat:=neBal ;
   InitCritEdt(Crit) ;
   Crit.Bal.TypCpt:=3 ; Crit.BAL.QuelAN:=AvecAN ;
   Crit.ModeRevision:=cbUnChecked ;
   Crit.Bal.FormatPrint.PrSepCompte:=False ;
   BalanceAuxiLegal(Crit) ;
   END ;
if (BalGen.Checked) and (Continuer) then
   BEGIN
   Crit.NatureEtat:=neBal ;
   InitCritEdt(Crit) ;
   Crit.Bal.TypCpt:=3 ; Crit.BAL.QuelAN:=AvecAN ;
   Crit.ModeRevision:=cbUnChecked ;
   Crit.Bal.FormatPrint.PrSepCompte:=False ;
   BalanceGeneLegal(Crit) ;
   END ;
if (JalDiv.Checked) and (Continuer) then
   BEGIN
   Crit.NatureEtat:=neJal ;
   InitCritEdt(Crit) ;
   Crit.ModeRevision:=cbUnChecked ;
   JalDivisioLegal(Crit) ;
   END ;
V_PGI.QRCouleur:=OrigineColor ;
if (GlvAux.Checked and GlvGen.Checked and BalAux.Checked
    and BalGen.Checked and JalDiv.Checked and Continuer)
    then PourClotureOk.State:=cbChecked Else PourClotureOk.State:=cbUnChecked ;
end;

procedure TEdLegal.BValiderClick(Sender: TObject);
begin
EnableControls(Self,False) ;
Continuer:=True ;
if CritEdtOk and PeriodesValidees then
   BEGIN
   //BStop.Enabled:=True ;
   LanceEdit ;
   //BStop.Enabled:=False ;
   END ;
EnableControls(Self,True) ;
end;

function  TEdLegal.CritEdtOk : Boolean ;
BEGIN
//Result:=True ;
Fillchar(Crit,SizeOf(Crit),#0) ;
With Crit Do
  BEGIN
  Exo.Code:=FExercice.Value ;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=DAte1 ; DateFin:=DAte2 ;
  if Not(GLvAux.Checked Or GLvGen.Checked
     Or BalAux.Checked Or BalGen.Checked Or JalDiv.Checked) then
     BEGIN
     MsgBleme.Execute(4,'','') ;
     Result:=False ;
     END else
     BEGIN
     Result:=CtrlPerExo(DateDeb, DateFin) ;
     if Not Result then MsgBleme.Execute(1,'','') ;
     END ;
  END ;
END ;

Function TEdLegal.PeriodesValidees : Boolean ;
Var Q               : TQuery ;
    Periode         : string[24] ;
    i,Nb,Prem,Dern,
    An1,An2,Reponse : integer ;
    Termine         : Boolean ;
    NouvDate        : TDateTime ;
BEGIN
Result:=False ; Termine:=False ;
Q:=OpenSql('Select EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+Crit.Exo.Code+'"',True) ;
Periode:=Q.FindField('EX_VALIDEE').AsString ;
Ferme(Q) ;
Prem:=FDateCpta1.ItemIndex+1 ; An1:=DonneAnnee(Crit.Date1) ;
Dern:=FDateCpta2.ItemIndex+1 ; An2:=DonneAnnee(Crit.Date2) ;
Nb:=Dern-Prem+1 ;
i:=Prem;
While (i<=Prem+Nb-1) and (Not Termine) do
  BEGIN
  if Periode[i]<>'X' then Termine:=True else Inc(i) ;
  END ;
if Termine then
  BEGIN
  if (i-Prem=0) then
     BEGIN
     if i>12 then NouvDate:=FinDeMois(EncodeDate(An2,i-12,1)) else NouvDate:=FinDeMois(EncodeDate(An1,i,1)) ;
     MsgBleme.Execute(2,FormatDateTime('mmmm yyyy',NouvDate),'') ;
     END else
     BEGIN
     if i>12 then NouvDate:=FinDeMois(EncodeDate(An2,i-13,1)) else NouvDate:=FinDeMois(EncodeDate(An1,i-1,1)) ;
     Reponse:=MsgBleme.Execute(3,FormatDateTime('mmmm yyyy',NouvDate),'') ;
     Case Reponse of
       mrYes : BEGIN Crit.Date2:=NouvDate ; Crit.DateFin:=Crit.Date2 ; Result:=True ; END ;
       mrNo  : ;
      end ;
     END ;
  END else
  BEGIN Result:=True ; END ;
END ;

procedure TEdLegal.FDateCpta1Change(Sender: TObject);
begin
FDate1.ItemIndex:=FDateCpta1.ItemIndex ; FDate2.ItemIndex:=FDateCpta2.ItemIndex ;
ChangeDateCompta ;

end;

procedure TEdLegal.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;

end;

procedure TEdLegal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action:=caFree ;
end;

procedure TEdLegal.FormCreate(Sender: TObject);
begin
(*
HorzScrollBar.Range:=0 ;HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ;VertScrollBar.Visible:=FALSE ;
ClientHeight:=HPB.Top+HPB.Height ; ClientWidth:=HPB.Left+HPB.Width ;
*)
end;

procedure TEdLegal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

end.

