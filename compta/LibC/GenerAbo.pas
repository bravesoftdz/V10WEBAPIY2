unit GenerAbo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFNDEF EAGLCLIENT}
  DBGrids, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDBG,
  {$ELSE}
  UTOB, 
  {$ENDIF}
  Hqry,
  StdCtrls,
  Buttons,
  Mask,
  Hctrls,
  ExtCtrls,
  ComCtrls,
  Grids,
  uObjFiltres, // GCO - 07/06/2006 - FQ 18308
  hmsgbox,
  Ent1,
  ParamDBG,
  HEnt1,
  DelVisuE,
  {$IFDEF AFAIRE}
   ValPerio,
  {$ENDIF}
  SaisComm, SaisUtil, LettUtil, Menus, HStatus, AboUtil, HDB,
  HSysMenu,  HTB97, ed_tools, HPanel, UiUtil,CONTABON_TOM,CPRECONABO_TOF,
{$IFNDEF CCS3}
  TiersPayeur,
{$ENDIF}
  TimpFic,
  utilSoc ;

type
  TFGenerAbo = class(TForm)
    Pages: TPageControl;
    PCriteres: TTabSheet;
    TG_CONTRAT: THLabel;
    CB_CONTRAT: THCritMaskEdit;
    CB_CONTRAT_: THCritMaskEdit;
    TCB_RECONDUCTION: TLabel;
    CB_RECONDUCTION: THValComboBox;
    HLabel1: THLabel;
    GroupBox1: TGroupBox;
    TE_DATECOMPTABLE: THLabel;
    DATEDEBUT: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    DATEFIN: THCritMaskEdit;
    CB_COMPTABLE: TCheckBox;
    HGA: THMsgBox;
    GA: THGrid;
    POPS: TPopupMenu;
    PModif: TPanel;
    PFenAbo: TPanel;
    H_TitreCouv: TLabel;
    Label1: TLabel;
    BCValide: TToolbarButton97;
    BCAbandon: TToolbarButton97;
    H_CONTRAT: TLabel;
    Label2: TLabel;
    FComboDate: TComboBox;
    H_LIBELLE: TLabel;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;

    Dock971: TDock97;
    Dock972: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    PPied: TToolWindow97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BValide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BZoomAbo: TToolbarButton97;
    BListePIECES: TToolbarButton97;
    BModifGenere: TToolbarButton97;
    TCB_QUALIFPIECE: THLabel;
    CB_QUALIFPIECE: THValComboBox;
    FindMvt: THFindDialog;
    procedure BAnnulerClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CB_CONTRATChange(Sender: TObject);
    procedure DATEFINExit(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BZoomAboClick(Sender: TObject);
    procedure GADblClick(Sender: TObject);
    procedure GAMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValideClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure POPSPopup(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BListePIECESClick(Sender: TObject);
    procedure BModifGenereClick(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure BCValideClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure CB_QUALIFPIECEChange(Sender: TObject);
  private
    QAbo        : TQuery ;
    GX,GY,NbSel : integer ;
    Date1,Date2 : TDateTime ;
    TAbo,TPIECE : TList ;
    GeneOBM     : TOBM ;
    FNomFiltre  : String ;
    WMinX,WMinY : Integer ;
    FObjFiltre : TObjFiltre;
    InfoImp           : TInfoImport ;
    //QFiche            : TQFiche ;

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
// Recherches, Remplissage
    Function  WhereCritAbo : string ;
    procedure EnteteChanged ( Ent : boolean ) ;
    procedure RempliGrid ;
    Function  JePrends : boolean ;
// Selection, Grid
    procedure CocheDecoche ( Lig : integer ; Next : boolean ) ;
    function  BonneGene ( Lig : integer ) : Boolean ;
    function  VireEtoile ( St : String ) : String ;
    procedure ModifAGenerer ;
    procedure BougeLimite ;
    procedure ChoixLimite ( Lig : integer ) ;
// Génération
    procedure BatchGenereAbo ;
    procedure MAJFicheABO ( Contrat : String ; NbGene : integer ; LastGene : TDateTime ) ;
  public
   Manuel,FindFirst : Boolean ;
   procedure ApresChargementFiltre;
  end;

Procedure GenereAbonnements ( Manuel : boolean ) ;

implementation

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi,
  UlibEcriture ;

const cNomFiltre = 'GENERABO';

{$R *.DFM}



Procedure GenereAbonnements ( Manuel : boolean ) ;
Var X  : TFGenerAbo ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch'],False,'nrSaisieModif') then Exit ;
X:=TFGenerAbo.Create(Application) ;
X.Manuel:=Manuel ; X.FNomFiltre:='GENERABO' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrSaisieModif',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFGenerAbo.BAnnulerClick(Sender: TObject);
begin
  Close ;
  {b FP FQ15555}
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
  {e FP FQ15555}
end;

procedure TFGenerAbo.FormCreate(Sender: TObject);
var Composants : TControlFiltre;
begin
  TAbo   := TList.Create ;
  TPIECE := TList.Create ;
  FNomFiltre:='' ;
  WMinX:=Width ;
  WMinY:=Height ;

  // GCO -* 07/06/2006 - Gestion des filtres FQ 18308
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  FObjFiltre := TObjFiltre.Create(Composants, '');
  FObjFiltre.ApresChangementFiltre := ApresChargementFiltre;
end;

procedure TFGenerAbo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(FObjFiltre); // GCO - 17/06/2006 - FQ 18308
  VideListe(TAbo) ;
  TAbo.Free ;
  VideListe(TPiece) ; TPiece.Free ;
  PurgePopup(POPS) ;
  if Parent is THPanel then
  begin
    _Bloqueur('nrSaisieModif',False) ;
    Action:=caFree ;
  end;
end;

procedure TFGenerAbo.POPSPopup(Sender: TObject);
begin InitPopup(Self) ; end;

procedure TFGenerAbo.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFGenerAbo.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFGenerAbo.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindMvt.Execute ; end;

procedure TFGenerAbo.FindMvtFind(Sender: TObject);
begin Rechercher(GA,FindMvt,FindFirst) ; end;

procedure TFGenerAbo.BModifGenereClick(Sender: TObject);
begin if BModifGenere.Enabled then ModifAGenerer ; end;

procedure TFGenerAbo.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
PrintDBGrid(GA,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFGenerAbo.CB_CONTRATChange(Sender: TObject);
begin EnteteChanged(True) ; end;

function TFGenerAbo.VireEtoile ( St : String ) : String ;
BEGIN
While St[1]='*' do Delete(St,1,1) ;
While St[Length(St)]='*' do Delete(St,Length(St),1) ;
VireEtoile:=St ;
END ;

procedure TFGenerAbo.ModifAGenerer ;
Var Lig : integer ;
    O   : TOBM ;
BEGIN
if Not Manuel then Exit ;
Lig:=GA.Row ; if ((Lig<=0) or (Lig>=GA.RowCount-1)) then Exit ;
if Lig>TAbo.Count then Exit ;
O:=TOBM(TAbo[Lig-1]) ; if O=Nil then Exit ;
if Not EstSelect(GA,Lig) then BEGIN HGA.Execute(12,caption,'') ; Exit ; END ;
ChoixLimite(Lig) ;
END ;

procedure TFGenerAbo.FormShow(Sender: TObject);
begin
  CB_QUALIFPIECE.Value:='N' ;

  // GCO - 07/06/2006 - Gestion des filtres FQ 18308
  FObjFiltre.FFI_TABLE := cNomFiltre;
  FObjFiltre.Charger;

  if Manuel then
  begin
     Caption:=HGA.Mess[14] ;
     HelpContext:=7328000 ;
  end
  else
  begin
    Caption:=HGA.Mess[0] ;
    HelpContext:=7331000 ;
  end;

  Date1:=VH^.Encours.Deb ;
  if VH^.DateCloturePer>Date1 then
    Date1:=VH^.DateCloturePer ;

  Date2:=FinDeMois(V_PGI.DateEntree) ;
  if Date2<Date1 then
    Date2:=Date1 ;

  // FQ 10742 - Pour éviter de générer des écritures sur un exercice inexistant
  if Date2 > VH^.Suivant.Fin then
    if Date2 > VH^.EnCours.Fin then
      Date2 := VH^.Encours.Fin;

  DateDebut.Text:=DateToStr(Date1) ; DateFin.Text:=DateToStr(Date2) ;
  GA.RowCount:=2 ; GA.ListeParam:='GENEREABO' ; EnteteChanged(True) ;
  if Not Manuel then
    BModifGenere.Enabled:=False ;

  UpdateCaption(Self) ;

  ApresChargementFiltre;

  HMTrad.ResizeGridColumns(GA) ;
end;

procedure TFGenerAbo.EnteteChanged ( Ent : boolean ) ;
BEGIN
if ENT then
   BEGIN
   GA.Enabled:=False ; BValide.Enabled:=False ; BCherche.Enabled:=True ;
   END else
   BEGIN
   GA.Enabled:=True ; BValide.Enabled:=True ; {BCherche.Enabled:=False ;}
   END ;
END ;

procedure TFGenerAbo.DATEFINExit(Sender: TObject);
Var DD : TDateTime ;
    Err : integer ;
begin
if Not IsValidDate(DATEFIN.Text) then
   BEGIN
   HGA.Execute(2,caption,'') ;
   DATEFIN.Text:=DateToStr(Date2) ;
   END else
   BEGIN
   DD:=StrToDate(DATEFIN.Text) ; Err:=DateCorrecte(DD) ;
   if Err>0 then
      BEGIN
      HGA.Execute(2+Err,caption,'') ;
      DATEFIN.Text:=DateToStr(Date2) ;
      END else Date2:=DD ;
   END ;
end;

{======================== Recherche, Remplissage du Grid ============================}
Function TFGenerAbo.WhereCritAbo : string ;
BEGIN
(*StAbo:='' ;
for j:=0 to Pages.PageCount-1 do
    BEGIN
    P:=Pages.Pages[j] ;
    for i:=0 to P.ControlCount-1 do
        BEGIN
        C:=P.Controls[i] ;
        Nam:=C.Name ; if Nam[Length(Nam)]='_' then System.Delete(Nam,Length(Nam),1) ;
      //  QAbo.Control2Criteres(Nam,StAbo,C,P) ;
        END ;
    END ;
if Copy(StAbo,2,3)='AND' then System.Delete(StAbo,1,5) ; if Copy(StAbo,2,2)='OR' then System.Delete(StAbo,1,4) ;
WhereCritAbo:=StAbo ; *)
WhereCritAbo:=RecupWhereCritere(Pages) ;
END ;

Function TFGenerAbo.JePrends : boolean ;
Var NextGene,LastGene : TDateTime ;
    Sep,Arr           : String3 ;
BEGIN
{$IFDEF TT}
result:= true ; exit ;
{$ENDIF}
JePrends:=False ;
if QAbo.FindField('CB_GUIDE').AsString='' then Exit ;
if QAbo.FindField('CB_DATECONTRAT').AsDateTime>Date2 then Exit ;
LastGene:=QAbo.FindField('CB_DATEDERNGENERE').AsDateTime ; if LastGene>Date2 then Exit ;
if (QAbo.FindField('CB_DEJAGENERE').AsInteger>=QAbo.FindField('CB_NBREPETITION').AsInteger) and (QAbo.FindField('CB_RECONDUCTION').AsString<>'TAC') then Exit ;
Sep:=QAbo.FindField('CB_SEPAREPAR').AsString ;
Arr:=QAbo.FindField('CB_ARRONDI').AsString ;
//NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
if (LastGene= idate1900) then NextGene:=QAbo.FindField('CB_DATECONTRAT').AsDateTime
                         else NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
if NextGene>Date2 then Exit ;
if ((Not Manuel) and (NextGene<Date1)) then Exit ;
JePrends:=True ;
END ;

procedure TFGenerAbo.RempliGrid ;
Var O : TOBM ;
    i : integer ;
BEGIN
VideListe(TAbo) ; NbSel:=0 ;
While Not QAbo.EOF do
   BEGIN
   if JePrends then
      BEGIN
      O:=TOBM.Create(EcrAbo,'',True) ;
      // LG* 21/02/2002 modif pour le tobm
      O.ChargeMvtP( QAbo , 'CB' );
      TAbo.Add(O) ;
      END ;
   QAbo.Next ;
   END ;
GA.VidePile(True) ;
for i:=0 to TAbo.Count-1 do
    BEGIN
    O:=TOBM(TAbo[i]) ; ComCom1(GA,O) ;
    if Not Manuel then
       BEGIN
       GA.Cells[GA.ColCount-1,GA.RowCount-1]:='+' ; Inc(NbSel) ;
       // YMO 03/11/2005 Système de Flipselection et IsSelected
       //(FQ10738 + Couleur = false dans la Grid)
       GA.FlipSelection(GA.Rowcount-1);
       END
       else
       // YMO 04/11/2005 RAZ de la grille si nouvelle recherche (FQ 16908)
       IF GA.IsSelected(GA.RowCount-1) then
       BEGIN
       GA.Cells[GA.ColCount-1,GA.RowCount-1]:=' ' ;
       GA.FlipSelection(GA.Rowcount-1);
       END ;
    GA.RowCount:=GA.RowCount+1 ;
    END ;

END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/03/2004
Modifié le ... :   /  /    
Description .. : - 17/03/2004 - LG - passage en eAGL
Mots clefs ... : 
*****************************************************************}
procedure TFGenerAbo.BChercheClick(Sender: TObject);
begin
DateFinExit(Nil) ;
VideListe(TAbo) ; //NbSel:=0 ;
if assigned(QAbo) then QAbo.Free ;
QAbo:=OpenSQL('Select * from CONTABON left join GUIDE on CB_GUIDE=GU_GUIDE '+WhereCritAbo +' AND GU_TYPE="ABO"',true) ;
RempliGrid ;
Ferme(QAbo) ; QAbo := nil ;
EnteteChanged(False) ;
HMTrad.ResizeGridColumns(GA) ;
end;

procedure TFGenerAbo.BZoomAboClick(Sender: TObject);
Var Lig  : integer ;
    O    : TOBM ;
    Okok : boolean ;
    St   : String ;
begin
Okok:=True ; St:='' ;
Lig:=GA.Row ; if ((Lig<=0) or (Lig>=GA.RowCount-1)) then Okok:=False ;
if Lig>TAbo.Count then Okok:=False ;
if Okok then
   BEGIN
   O:=TOBM(TAbo[Lig-1]) ; if O=Nil then Exit ;
   St:=VireEtoile(O.GetMvt('CB_CONTRAT')) ;
   END ;
ParamAbonnement(True,St,taModif) ;
BChercheClick(nil);
NbSel:=0 ; 
end;

{================================ SELECTION =====================================}
function TFGenerAbo.BonneGene ( Lig : integer ) : Boolean ;
Var LastGene,NextGene,DateModif,DateCon : TDateTime ;
    SEP,ARR  : String3 ;
    O        : TOBM ;
    ii       : integer ;
    Q        : TQuery ;
BEGIN
Result:=True ; if Not Manuel then Exit ;
O:=TOBM(TAbo[Lig-1]) ; if O=Nil then Exit ;
LastGene:=O.GetMvt('CB_DATEDERNGENERE') ; Sep:=O.GetMvt('CB_SEPAREPAR') ; Arr:=O.GetMvt('CB_ARRONDI') ;
if (LastGene= idate1900) then NextGene:=O.GetMvt('CB_DATECONTRAT')
                         else NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
//NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
DateModif:=Date1 ; DateCon:=O.GetMvt('CB_DATECONTRAT') ; if DateModif<DateCon then DateModif:=DateCon ;
if NextGene<DateModif then
   BEGIN
   ii:=HGA.Execute(8,caption,'') ;
   if ii<>mrYes then Result:=False else
      BEGIN
      if Not ModifAboGene(VireEtoile(O.GetMvt('CB_CONTRAT')),DateModif) then
         BEGIN
         HGA.Execute(10,caption,'') ; Result:=False ;
         END else
         BEGIN
         Q:=OpenSQL('Select * from CONTABON Where CB_CONTRAT="'+VireEtoile(O.GetMvt('CB_CONTRAT'))+'"',True) ;
         if Not Q.EOF then BEGIN O.ChargeMvt(Q) ; ComComLigne(GA,O,Lig) ; END else Result:=False ;
         Ferme(Q) ;
         END ;
      END ;
   END ;
END ;

procedure TFGenerAbo.BCAbandonClick(Sender: TObject);
begin
Pages.Enabled:=True ; GA.Enabled:=True ; PPied.Enabled:=True ;
//vt if BModifGenere.CanFocus then BModifGenere.SetFocus ;
PModif.Visible:=False ;
end;

procedure TFGenerAbo.BCValideClick(Sender: TObject);
Var NbChoix : integer ;
    O       : TOBM ;
    St      : String ;
begin
NbChoix:=FComboDate.ItemIndex+1 ;
O:=TOBM(TAbo[GA.Row-1]) ; if O=Nil then Exit ;
O.NbDiv:=NbChoix ;
St:=O.GetMvt('CB_CONTRAT') ; St:=VireEtoile(St) ;
if NbChoix<FComboDate.Items.Count then St:='**'+St+'**' ;
O.PutMvt('CB_CONTRAT',St) ; ComComLigne(GA,O,GA.Row) ;
BCAbandonClick(Nil) ;
end;

procedure TFGenerAbo.BougeLimite ;
BEGIN
PModif.Left:=GA.Left+(GA.Width-PModif.Width) div 2 ;
PModif.Top:=GA.Top+(GA.Height-PModif.Height) div 2 ;
END ;

procedure TFGenerAbo.ChoixLimite ( Lig : integer ) ;
Var O : TOBM ;
    LastGene,NextGene : TDateTime ;
    SEP,ARR,REC : String3 ;
    NbRepet,NbGene,NbChoix : integer ;
BEGIN
O:=TOBM(TAbo[Lig-1]) ; NbChoix:=O.NbDiv ;
FComboDate.Items.Clear ;
LastGene:=O.GetMvt('CB_DATEDERNGENERE') ;
NbGene:=O.GetMvt('CB_DEJAGENERE') ; NbRepet:=O.GetMvt('CB_NBREPETITION') ;
SEP:=O.GetMvt('CB_SEPAREPAR') ; ARR:=O.GetMvt('CB_ARRONDI') ; REC:=O.GetMvt('CB_RECONDUCTION') ;
NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
Repeat
 LastGene:=NextGene ; FComboDate.Items.Add(DateToStr(LastGene)) ;
 NextGene:=ProchaineDate(LastGene,SEP,ARR) ; NbGene:=NbGene+1 ;
 if NextGene<Date1 then Break ;
 if NextGene>Date2 then Break ;
 if NextGene<=LastGene then Break ;
 if REC='TAC' then
    BEGIN
    if NbGene>=NbRepet then NbGene:=0 ;
    END else
    BEGIN
    if NbGene>=NbRepet then Break ;
    END ;
Until False ;
H_CONTRAT.Caption:=VireEtoile(O.GetMvt('CB_CONTRAT')) ;
H_LIBELLE.Caption:=VireEtoile(O.GetMvt('CB_LIBELLE')) ;
if NbChoix<=0 then FComboDate.ItemIndex:=FComboDate.Items.Count-1 else FComboDate.ItemIndex:=NbChoix-1 ;
BougeLimite ;
PModif.Visible:=True ; FComboDate.SetFocus ;
Pages.Enabled:=False ; GA.Enabled:=False ; PPied.Enabled:=False ;
END ;

procedure TFGenerAbo.CocheDecoche ( Lig : integer ; Next : boolean ) ;
BEGIN
if Lig>=GA.RowCount-1 then Exit ;
if EstSelect(GA,Lig) then
   BEGIN
   GA.Cells[GA.ColCount-1,Lig]:=' ' ; Dec(NbSel) ;
   END else
   BEGIN
   if Not BonneGene(Lig) then Exit ;
   GA.Cells[GA.ColCount-1,Lig]:='+' ; Inc(NbSel) ;
   END ;
if Next then BEGIN GA.Row:=Lig+1 ; GA.Invalidate ; END ;
END ;

{================================ METHODES GRID, FORM =====================================}
procedure TFGenerAbo.GADblClick(Sender: TObject);
Var C,R : Longint ;
begin
GA.MouseToCell(GX,GY,C,R) ;
if ((R>0) and (R<GA.RowCount-1)) then CocheDecoche(GA.Row,True) ;
end;

procedure TFGenerAbo.GAMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   GA.MouseToCell(X,Y,C,R) ;
   if ((R>0) and (R<GA.RowCount-1)) then CocheDecoche(GA.Row,False) ;
   END ;
end;

procedure TFGenerAbo.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
begin
OkG:=GA.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_F9     : BEGIN Key:=0 ; if Vide then BChercheClick(Nil) ; END ;
   VK_F10    : BEGIN Key:=0 ; if Vide then BValideClick(Nil) ; END ;
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(GA.Row,False) ;
//13575   VK_F5     : BEGIN Key:=0 ; BZoomAboClick(Nil) ; END ;
{^F}      70 : BEGIN Key:=0 ; if Shift=[ssCtrl] then BChercheClick(Nil) ; END ;
{^L}      76 : BEGIN Key:=0 ; if Shift=[ssCtrl] then BListePiecesClick(Nil) ; END ;
{AM}      77 : BEGIN Key:=0 ; if Shift=[ssAlt] then BModifGenereClick(Nil) ; END ;
{^P}      80 : BEGIN Key:=0 ; if Shift=[ssCtrl] then BImprimerClick(Nil) ; END ;
  END ;
end ;

{================================ GENERATION =====================================}
procedure TFGenerAbo.MAJFicheABO ( Contrat : String ; NbGene : integer ; LastGene : TDateTime ) ;
BEGIN
Contrat:=VireEtoile(Contrat) ;
if ExecuteSQL('UPDATE CONTABON Set CB_DEJAGENERE='+IntToStr(NbGene)+', CB_DATEDERNGENERE="'+USDATETIME(LastGene)+'" Where CB_COMPTABLE="X" and CB_CONTRAT="'+Contrat+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 05/09/2007
Modifié le ... :   /  /    
Description .. : LG - 05/09/2007 - coorection d ela generation des abo en 
Suite ........ : tva sur encaissement
Mots clefs ... : 
*****************************************************************}
procedure TFGenerAbo.BatchGenereAbo ;
Var CodeG  : String ;
    ContratG,LibContratG : String ;
    Q      : TQuery ;
    M      : RMVT ;
    LastGene,NextGene : TDateTime ;
    NbGene,NbRepet,NbChoix,NbIter : integer ;
    SEP,ARR,REC       : String3 ;
    DEV               : RDEVISE ;
    Trouv,Okok        : Boolean ;
    O                 : TOBM ;
    DD                : TDateTime ;
BEGIN
{Entête de guide}
try
O:=GeneOBM ; CodeG:=O.GetMvt('CB_GUIDE') ;
ContratG:=O.GetMvt('CB_CONTRAT') ; LibContratG:=O.GetMvt('CB_LIBELLE') ;
Q:=OpenSQL('Select * from GUIDE Where GU_TYPE="ABO" and GU_GUIDE="'+CodeG+'"',True) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   FillChar(M,Sizeof(M),#0) ;
   M.Etabl:=Q.FindField('GU_ETABLISSEMENT').AsString ;
   M.Jal:=Q.FindField('GU_JOURNAL').AsString ;
   M.CodeD:=Q.FindField('GU_DEVISE').AsString ; if M.CodeD='' then M.CodeD:=V_PGI.DevisePivot ;
   M.Valide:=False ; M.EtapeRegle:=False ;
   M.Nature:=Q.FindField('GU_NATUREPIECE').AsString ;
   M.Simul:=O.GetMvt('CB_QUALIFPIECE') ;
   END ;
Ferme(Q) ; if Not Trouv then Exit ;
{Paramètres de génération}
LastGene:=O.GetMvt('CB_DATEDERNGENERE') ;
NbGene:=O.GetMvt('CB_DEJAGENERE') ; NbRepet:=O.GetMvt('CB_NBREPETITION') ;
SEP:=O.GetMvt('CB_SEPAREPAR') ; ARR:=O.GetMvt('CB_ARRONDI') ; REC:=O.GetMvt('CB_RECONDUCTION') ;
if (LastGene= idate1900) then NextGene:=O.GetMvt('CB_DATECONTRAT')
                         else NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
//NextGene:=ProchaineDate(LastGene,SEP,ARR) ;
M.DateC:=NextGene ; M.Exo:=QuelExoDT(M.DateC) ;
FillChar(DEV,Sizeof(DEV),#0) ; Okok:=False ;
DEV.Code:=M.CodeD ; GetInfosDevise(DEV) ;
DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,M.DateC) ;
M.TauxD:=DEV.Taux ; M.DateTaux:=DEV.DateTaux ;
{Génération écritures}
NbIter:=0 ; NbChoix:=O.NbDiv ; DD:=0 ;
Repeat
 If M.Simul='N' Then DD:=M.DateC ;
 M.Num:=GetNewNumJal(M.Jal,(M.Simul='N'),DD) ; if M.Num<=0 then Break ;
 EcrituresAbo(CodeG,M,DEV,TPIECE,ContratG,LibContratG) ; Okok:=True ; Inc(NbIter) ;
 if M.Simul='N' then
    BEGIN

    //MajSoldesPiece(M) ;
    ADevalider(M.Jal,M.DateC) ;
    {#TVAENC}
    if VH^.OuiTvaEnc then //ElementsTvaEnc(M,False) ;
  //  ElementsTvaEncRevise (M, False, InfoImp, QFiche) ;   // FB 21156 -  ne plus utiliser cette fct
    ElementsTvaEnc(M,False) ;
{$IFNDEF CCS3}
    if VH^.OuiTP then GenerePiecesPayeur(M) ;
{$ENDIF}
    END ;
 LastGene:=NextGene ; NextGene:=ProchaineDate(LastGene,SEP,ARR) ; NbGene:=NbGene+1 ;
 if NextGene<Date1 then Break ;
 if NextGene>Date2 then Break ;
 if NextGene<=LastGene then Break ;
 if ((NbChoix>0) and (NbIter>=NbChoix)) then Break ;
 if REC='TAC' then
    BEGIN
    if NbGene>=NbRepet then NbGene:=0 ;
    END else
    BEGIN
    if NbGene>=NbRepet then Break ;
    END ;
 M.DateC:=NextGene ; M.Exo:=QuelExoDT(M.DateC) ;
Until False ;
if Okok then
   BEGIN
   MajFicheABO(VireEtoile(O.GetMvt('CB_CONTRAT')),NbGene,LastGene) ;
   MarquerPublifi(True) ;
   CPStatutDossier ;
   END ;
except
 on E : exception do
  begin
   V_PGI.IOError := oeUnknown ;
   PGIError('Erreur lors de la génération des abonnements !' + #10#13 + e.message ) ;
  end ;
end ;
END ;

procedure TFGenerAbo.BValideClick(Sender: TObject);
Var i : integer ;
    O : TOBM ;
    Erreur : TIOErr ;
begin
if NbSel<=0 then BEGIN HGA.Execute(7,caption,'') ; Exit ; END ;
if HGA.Execute(9,caption,'')<>mrYes then Exit ;
if RevisionActive(StrToDate(DateDebut.Text)) then Exit ;
Application.ProcessMessages ;
InitMove(NbSel,'') ;
for i:=1 to GA.RowCount-2 do if EstSelect(GA,i) then
BEGIN
   O:=TOBM(TAbo[i-1]) ; GeneOBM:=O ; MoveCur(FALSE) ;
   { FQ 20001 BVE 25.04.07 }
   Erreur := Transactions(BatchGenereAbo,0);
   if Erreur <> oeOK then
   begin
      if Erreur = oeLettrage then
         MessageAlerte(HGA.Mess[17])  // Erreur cpt Général
      else if Erreur = oeTiers then
         MessageAlerte(HGA.Mess[18])  // Erreur cpt Auxiliaire
      else if Erreur = oeStock then
         MessageAlerte(HGA.Mess[19]) ; // Erreur Journal
  //    else
  //       MessageAlerte(HGA.Mess[13]) ; // Autre
      FiniMove ;
      Exit ;
   end ;
   { END FQ 20001 BVE 25.04.07 }
END ;
FiniMove ;
NbSel:=0 ; 
BChercheClick(Nil) ;
if HGA.Execute(11,caption,'')=mrYes then
 VisuPiecesGenere(TPiece,EcrGen,7) ;
end;

procedure TFGenerAbo.BListePIECESClick(Sender: TObject);
begin
if TPiece.Count>0 then VisuPiecesGenere(TPiece,EcrGen,7) ;
end;

procedure TFGenerAbo.FormResize(Sender: TObject);
begin BougeLimite ; end;

procedure TFGenerAbo.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFGenerAbo.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFGenerAbo.CB_QUALIFPIECEChange(Sender: TObject);
begin
if ((CB_QUALIFPIECE.Value='R') and (Not V_PGI.Controleur)) then
   BEGIN
   HGA.Execute(16,Caption,'') ;
   CB_QUALIFPIECE.Value:='N' ;
   END ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFGenerAbo.ApresChargementFiltre;
begin
  if FFiltres.Text <> '' then
    BChercheClick(BCherche)
  else
    CB_RECONDUCTION.ItemIndex := 0;
end;

////////////////////////////////////////////////////////////////////////////////
end.
