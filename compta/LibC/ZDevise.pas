unit ZDevise;

//V_PGI.DeviseFongible
//VH^.LibDeviseFongible
//V_PGI.TauxEuro

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Hctrls, StdCtrls, Mask, Buttons, ExtCtrls,
  HEnt1, UTob,
  Ent1, SaisUtil,
  CPCHANCELL_TOF,     // pour FicheChancel
  DEVISE_TOM,
  SaisTaux1,
  TZ, HSysMenu, hmsgbox ;

type RSDev = record
             Ecr                        : TOB ;
             DPivot                     : RDevise ;
             TD, TC, Solde              : Double;
             SensDef                    : string ;
             bMonoDev                   : boolean ;
             Action                     : TActionFiche ;
             bDoSolde                   : Boolean ;
             DeviseTiers                : string ;
             DeviseLigne                : RDevise ;
             volatile                   : boolean
             end ;

function SaisieZDevise( var Params : RSDev) : Boolean ;

const RC_VERIFDEVISE     =  0 ;
      RC_VERIFMONTANT    =  1 ;
      RC_VALIDER         =  2 ;
      RC_PASNEGATIF      =  3 ;
      RC_HORSFOURCHETTE  =  4 ;
      RC_PARITEFAUSSE    =  5 ;
      RC_TAUXA1          =  6 ;
      RC_VOLATILPERDU    =  7 ;
      RC_NOPARITEFIXE    =  8 ;
      RC_SAISIRTAUX      =  9 ;



type
  TZFDevise = class(TForm)
    Panel1: TPanel;
    HLabel2: THLabel;
    FDate: TMaskEdit;
    HLDEVISE: THLabel;
    FDevise: THValComboBox;
    HLTAUX: THLabel;
    FTaux: THNumEdit;
    HLDEBIT: THLabel;
    HLMONTANTD: THLabel;
    FPivotDebit: THNumEdit;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    HLCREDIT: THLabel;
    FPivotCredit: THNumEdit;
    HLMONTANTC: THLabel;
    DockTop: TDock97;
    DockLeft: TDock97;
    DockRight: TDock97;
    DockBottom: TDock97;
    Outils: TToolbar97;
    BChancel: TToolbarButton97;
    BSolde: TToolbarButton97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    SepBevel: TBevel;
    HLTDEBIT: THLabel;
    HLTCREDIT: THLabel;
    HLTSOLDE: THLabel;
    FTotalSolde: THNumEdit;
    FTotalDebit: THNumEdit;
    FTotalCredit: THNumEdit;
    DEB: TEdit;
    CRE: TEdit;
    BSaisTaux: TToolbarButton97;
    ISigneEuro: TImage;
    procedure BChancelClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure FDeviseChange(Sender: TObject);
    procedure FEuroClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BSaisTauxClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FEuroExit(Sender: TObject);
    procedure DEBChange(Sender: TObject);
    procedure CREChange(Sender: TObject);
  private
    memDevise : RDevise ;
    bVolatile : Boolean ;
    Ecr      : TOB ;
    Pivot    : RDevise ;
    TD, TC   : Double ; // Cumul Débit, Crédit de la pièce
    ED, EC   : Double ; // Débit, Crédit de la ligne avant modification
    ESolde   : Double ; // Solde calculé de la pièce avant modification
    SensDef  : string ; // Sens par défaut du compte en cours de saisie
    bMono    : Boolean ; // Mono devise pour le tiers
    bDoSolde : Boolean ; // Faire le solde en premier lieu
    bLoad    : Boolean ; // initialisation

    // Fonctions ressource
    function  PrintMessageRC(MessageID : Integer; sCaption : string='' ; StAfter: string='') : Integer ;
    procedure EnableButtons ;
    procedure Calcul ;
    procedure InitEntete ;
    procedure InitFocus ;
    function PutMontant(Montant : Double) : string ;
    function  ControleSaisie(bVerbose : Boolean) : Boolean ;
    procedure SoldeClick ;
    procedure TauxClick ;
  public
    Action   : TActionFiche ;
    constructor Create(AOwner: TComponent; Params: RSDev) ; reintroduce ;
    function    PbTaux(Dev: RDevise; DateCpt: TDateTime): Boolean ;
    procedure   AvertirPbTaux(Code: String3; DateTaux, DateECr: TDateTime; bForce: Boolean=FALSE) ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  uLibEcriture ;


//=======================================================
//======== Point d'entrée dans la saisie devise =========
//=======================================================
function CGetCotation(DateEcr : TDateTime; Taux : Double; Dev : String3) : Double ;
var Cote : Double ;
begin
if DateEcr<V_PGI.DateDebutEuro then Cote:=Taux else
   begin
   if Dev=V_PGI.DevisePivot then Cote:=1.0 else
     if V_PGI.TauxEuro<>0 then Cote:=V_PGI.TauxEuro/Taux else Cote:=1 ;
   end ;
Result:=Cote ;
end ;

//=======================================================
//======== Point d'entrée dans la saisie devise =========
//=======================================================
function SaisieZDevise(var Params : RSDev) : Boolean ;
var ZFDevise : TZFDevise ;
begin
ZFDevise:=TZFDevise.Create(Application, Params) ;
  try
  Result:=(ZFDevise.ShowModal=mrOK) ;
  Params.DeviseLigne := ZFDevise.memDevise ;
  Params.volatile    := ZFDevise.bVolatile ;
  finally
  ZFDevise.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

constructor TZFDevise.Create(AOwner: TComponent; Params: RSDev) ;
begin
Ecr:=Params.Ecr ;
Pivot:=Params.DPivot ;
ESolde:=Params.Solde ;
TD:=Params.TD ;
TC:=Params.TC ;
SensDef:=Params.SensDef ;
bMono:=Params.bMonoDev ;
Action:=Params.Action ;
bDoSolde:=Params.bDoSolde ;
bVolatile := False ;
if Params.volatile then
  begin
  bVolatile := True ;
  memDevise := Params.DeviseLigne ;
  end ;
inherited Create(AOwner) ;
end ;


//=======================================================
//================= Fonctions Ressource =================
//=======================================================

function TZFDevise.PrintMessageRC(MessageID : Integer; sCaption : string ; StAfter : string) : Integer ;
begin
if sCaption<>'' then Result:=HM.Execute(MessageID, sCaption, StAfter)
                else Result:=HM.Execute(MessageID, Caption,  StAfter) ;
end ;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
function TZFDevise.PutMontant(Montant : Double) : string ;
var Dec : Integer ;
begin
Dec:=memDevise.Decimale ;
Result:=StrFMontant(Montant, 15, Dec, '', TRUE) ;
end ;

function TZFDevise.ControleSaisie(bVerbose : Boolean) : Boolean ;
begin
Result:=TRUE ;
if (FDevise.Value='') then
   begin if bVerbose then HM.Execute(0,'','') ; Result:=FALSE ; Exit ; end ;
if (Valeur(DEB.Text)=0) and (Valeur(CRE.Text)=0) then
   begin if bVerbose then HM.Execute(1,'','') ; Result:=FALSE ; Exit ; end ;
if (not VH^.MontantNegatif) and ((FPivotDebit.Value<0) or (FPivotCredit.Value<0)) then
   begin if bVerbose then HM.Execute(3,'','') ; Result:=FALSE ; Exit ; end ;
if ((FPivotDebit.Value<>0) and (FPivotDebit.Value<VH^.GrpMontantMin) or (FPivotDebit.Value>VH^.GrpMontantMax)) then
   begin if bVerbose then HM.Execute(4,'','') ; Result:=FALSE ; Exit ; end ;
if ((FPivotCredit.Value<>0) and (FPivotCredit.Value<VH^.GrpMontantMin) or (FPivotCredit.Value>VH^.GrpMontantMax)) then
   begin if bVerbose then HM.Execute(4,'','') ; Result:=FALSE ; Exit ; end ;
end ;


procedure TZFDevise.EnableButtons ;
begin
if Action<>taConsult then
  begin
  DEB.Enabled:=TRUE ;
  CRE.Enabled:=TRUE ;
  if Valeur(DEB.Text)<>0 then
    CRE.Enabled:=FALSE ;
  if Valeur(CRE.Text)<>0 then
    DEB.Enabled:=FALSE ;
  BSolde.Enabled:=(FTotalSolde.Value<>0) ;
  end
else
  begin
  FDevise.Enabled := FALSE ;
  DEB.Enabled     := FALSE ;
  CRE.Enabled     := FALSE ;
  BSolde.Enabled  := FALSE ;
  end ;
end ;

procedure TZFDevise.Calcul ;
begin
CSetMontants(Ecr, Valeur(DEB.Text), Valeur(CRE.Text), memDevise,TRUE) ;
FPivotDebit.Value:=Ecr.GetValue('E_DEBIT') ;
FPivotCredit.Value:=Ecr.GetValue('E_CREDIT') ;
AfficheLeSolde(FTotalSolde, (TD-ED)+Valeur(DEB.Text), (TC-EC)+Valeur(CRE.Text)) ;
EnableButtons ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 07/11/2006
Modifié le ... : 07/11/2006
Description .. : - LG - 07/11/2006 - Suite demande GPAO, on permet de 
Suite ........ : saisir une devise par défaut en multi devises
Mots clefs ... : 
*****************************************************************}
procedure TZFDevise.InitEntete ;
begin
bLoad := True ;
BSaisTaux.Visible:=(Action=taCreat) ;
//bVolatile:=FALSE ;
if Ecr.GetValue('E_DEVISE')<>Pivot.Code then
   begin
   memDevise.Code:=Ecr.GetValue('E_DEVISE') ;
   if not bvolatile then
     begin
     GetInfosDevise(memDevise) ;
     memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
     end ;
   end else
   begin
   memDevise.Code:=Pivot.Code ;
   GetInfosDevise(memDevise) ;
   memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, Ecr.GetValue('E_DATECOMPTABLE')) ;
   end ;
   DEB.Text:=StrFMontant(Ecr.GetValue('E_DEBITDEV'),15,memDevise.Decimale,'',TRUE) ;
   CRE.Text:=StrFMontant(Ecr.GetValue('E_CREDITDEV'),15,memDevise.Decimale,'',TRUE) ;
if Ecr.GetValue('E_DEVISE')<>Pivot.Code then
   FDevise.Value:=Ecr.GetValue('E_DEVISE') ;

  FDevise.Enabled := not bMono and (Action=taCreat);

FDate.Text:=DateToStr(Ecr.GetValue('E_DATECOMPTABLE')) ;
FPivotDebit.Value:=Ecr.GetValue('E_DEBIT') ;
FPivotCredit.Value:=Ecr.GetValue('E_CREDIT') ;
FTotalDebit.Value:=TD ;  ED:=Valeur(DEB.Text) ;
FTotalCredit.Value:=TC ; EC:=Valeur(CRE.Text) ;
FTotalSolde.Value:=ESolde ;
HLMONTANTD.Caption:=Pivot.Code ;
HLMONTANTC.Caption:=Pivot.Code ;
Calcul ;
bLoad := False ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 17/12/2002
Modifié le ... :   /  /    
Description .. : 17/12/2002 - fiche 10833 - correction de la gestion du 
Suite ........ : focus
Mots clefs ... : 
*****************************************************************}
procedure TZFDevise.InitFocus ;
var ValD, ValC : Double ;
begin
  ValD:=Valeur(DEB.Text) ;
  ValC:=Valeur(CRE.Text) ;
  if Action<>taConsult then
    begin
    if (ValD+ValC=0) and FDevise.Enabled then
      begin
      FDevise.SetFocus ;
      exit ;
      end ;
  //   if (SensDef='D') or (SensDef='M') then DEB.SetFocus else CRE.SetFocus ;
    // CA - 05/02/2003 : Focus fait main : attention SensDef est toujours à C ( à voir ...)
    // Montant déja saisi ==> prioritaire
    if ValD<>0
      then DEB.SetFocus
      else if valC <> 0
        then CRE.SetFocus
    // Montant à 0, init par défaut :
    // --> soit sur reste à solder...
    else if ESolde > 0 
      then CRE.SetFocus
      else if ESolde < 0
        then DEB.SetFocus
    // --> soit sur paramètre d'appel (sens du compte, nature piece...)
    else if (SensDef='D') or (SensDef='M')
      then DEB.SetFocus
      else CRE.SetFocus ;
    end ;
end ;

procedure TZFDevise.SoldeClick ;
var  Debit, Credit, Solde : Double ; bSoldeDebit : Boolean ; Dec : Integer ;
begin
if not BSolde.Enabled then Exit ;
Debit:=Valeur(DEB.Text) ; Credit:=Valeur(CRE.Text) ;
Dec:=memDevise.Decimale ;
Solde:=Arrondi(((TD-ED)+Debit)-((TC-EC)+Credit), Dec) ;
bSoldeDebit:=(Solde>0) ;
Solde:=Abs(Solde) ;
DEB.Text:=PutMontant(0) ;
CRE.Text:=PutMontant(0) ;
// Partie débit
if (Debit<>0) and (bSoldeDebit) then
   if (Debit-Solde)<0 then CRE.Text:=PutMontant(Abs(Debit-Solde))
                      else DEB.Text:=PutMontant(Abs(Debit-Solde)) ;
if (Debit<>0) and (not bSoldeDebit) then
   DEB.Text:=PutMontant(Abs(Debit+Solde)) ;
// Partie crédit
if (Credit<>0) and (bSoldeDebit) then
   CRE.Text:=PutMontant(Abs(Credit+Solde)) ;
if (Credit<>0) and (not bSoldeDebit) then
   if (Credit-Solde)<0 then DEB.Text:=PutMontant(Abs(Credit-Solde))
                       else CRE.Text:=PutMontant(Abs(Credit-Solde)) ;
if (Debit=0) and (Credit=0) then
   begin
   if bSoldeDebit then CRE.Text:=PutMontant(Solde)
                  else DEB.Text:=PutMontant(Solde) ;
   end ;
Calcul ;
if Valeur(DEB.Text)<>0 then DEB.SetFocus else CRE.SetFocus ;
EnableButtons ;
end ;

procedure TZFDevise.TauxClick ;
var memTmp : RDevise ;
begin
if not BSaisTaux.Enabled then Exit ;
if memDevise.Code=V_PGI.DevisePivot then Exit ;
if memDevise.Code='' then Exit ;
if Action<>taCreat then Exit ;
memTmp:=memDevise ;
if SaisieNewTaux2000(memTmp, Ecr.GetValue('E_DATECOMPTABLE')) then
  begin
  memDevise.Taux:=memTmp.Taux ;
  FTaux.Value:=CGetCotation(Ecr.GetValue('E_DATECOMPTABLE'), memDevise.Taux, memDevise.Code) ;
  bVolatile:=TRUE ; Calcul ;
  end ;
end ;

function TZFDevise.PbTaux(Dev: RDevise; DateCpt: TDateTime): Boolean ;
var Code: string3 ;
begin
Result:=FALSE ;
Code:=Dev.Code ;
if ((Code=V_PGI.DevisePivot) or (Code=V_PGI.DeviseFongible)) then Exit ;

if ((DateCpt<V_PGI.DateDebutEuro) or (Code<>V_PGI.DevisePivot) {(Not EstMonnaieIn(Code))}) then Result:=(DEV.Taux=1)
//else if EstMonnaieIn(Code) then Result:=(DEV.Taux=V_PGI.TauxEuro) ;
end ;

procedure TZFDevise.AvertirPbTaux(Code: String3; DateTaux, DateECr: TDateTime; bForce: Boolean=FALSE) ;
var i: Integer ; bTauxOk: Boolean ;
begin
if (not bForce) and (Action<>taCreat) then Exit ;
bTauxOk:=(Arrondi(memDevise.Taux-1, ADecimP)=0) ;
(*
if ((EstMonnaieIN(Code)) and (DateEcr>=V_PGI.DateDebutEuro)) then
  begin
  i:=PrintMessageRC(RC_NOPARITEFIXE) ;
  if i<>mrYes then i:=PrintMessageRC(RC_PARITEFAUSSE) ;
  if i=mrYes then
    begin
    FicheDevise(Code, taModif, FALSE) ;
    memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, DateEcr) ;
    end ;
  end else
  *)
  begin
  if bTauxOk then i:=PrintMessageRC(RC_TAUXA1) else i:=PrintMessageRC(RC_SAISIRTAUX) ;
  if i=mrYes then
    begin
    FicheChancel(FDevise.Value, TRUE, DateEcr, taCreat, (DateEcr>=V_PGI.DateDebutEuro)) ;
    memDevise.Taux:=GetTaux(memDevise.Code, memDevise.DateTaux, DateEcr) ;
    end ;
  end ;

end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TZFDevise.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if ModalResult<>mrOk then
  begin
  if ControleSaisie(FALSE) then
    begin
      Ecr.PutValue('E_DEVISE',     memDevise.Code) ;
      Ecr.PutValue('E_TAUXDEV',    memDevise.Taux) ;
      ModalResult:=mrOk
    end
    else ModalResult:=mrCancel ;
  end ;
end ;

procedure TZFDevise.FormShow(Sender: TObject);
begin
EnableButtons ;
if bDoSolde then BSoldeClick(nil) ;
InitFocus ;
end;

//=======================================================
//=============== Evénements des boutons ================
//=======================================================
procedure TZFDevise.BChancelClick(Sender: TObject);
var DateEcr : TDateTime ; ChAction : TActionFiche ;
begin
DateEcr:=Ecr.GetValue('E_DATECOMPTABLE') ;
if Action<>taConsult then
  begin
  memDevise.Taux:=GetTaux(FDevise.Value, memDevise.DateTaux, DateEcr) ;
  if memDevise.DateTaux<>DateEcr then ChAction:=taCreat else ChAction:=taModif ;
  end else ChAction:=Action ;
(*
if EstMonnaieIn(FDevise.Value)
  then FicheDevise(FDevise.Value, ChAction, FALSE)
  else *)FicheChancel(FDevise.Value, TRUE, DateEcr, ChAction, (DateEcr>=V_PGI.DateDebutEuro)) ;
if Action<>taConsult then
  begin
  memDevise.Taux:=GetTaux(FDevise.Value, memDevise.DateTaux, DateEcr) ;
    FTaux.Value:=CGetCotation(DateEcr, memDevise.Taux, memDevise.Code) ;
  Calcul ;
  end ;
end ;


procedure TZFDevise.BValideClick(Sender: TObject);
begin
if Action<>taConsult then
  begin
  if not ControleSaisie(TRUE) then Exit ;
  Calcul ;
  Ecr.PutValue('E_DEVISE',     memDevise.Code) ;
  Ecr.PutValue('E_TAUXDEV',    memDevise.Taux) ;
  end ;
ModalResult:=mrOk ;
end;

procedure TZFDevise.FDeviseChange(Sender: TObject);
var DateEcr: TDateTime ;
begin
if (not bLoad) and bVolatile then begin HM.Execute(7,'','') ; bVolatile:=FALSE ; end ;
DateEcr:=Ecr.GetValue('E_DATECOMPTABLE') ;
if not bVolatile then
  begin
  memDevise.Code:=FDevise.Value ;
  GetInfosDevise(memDevise) ;
  memDevise.Taux:=GetTaux(FDevise.Value, memDevise.DateTaux, DateEcr) ;
  if (memDevise.DateTaux<>DateEcr) or PbTaux(memDevise, DateEcr)
    then AvertirPbTaux(memDevise.Code, memDevise.DateTaux, DateEcr) ;
  end ;
FTaux.Value:=CGetCotation(DateEcr, memDevise.Taux, memDevise.Code) ;
Calcul ;
ChangeMask(FTotalDebit,  memDevise.Decimale, memDevise.Symbole) ;
ChangeMask(FTotalCredit, memDevise.Decimale, memDevise.Symbole) ;
ChangeMask(FTotalSolde,  memDevise.Decimale, memDevise.Symbole) ;
end;

procedure TZFDevise.FEuroClick(Sender: TObject);
begin
if bMono then FDevise.Enabled:=FALSE ;
FTaux.Value:=CGetCotation(Ecr.GetValue('E_DATECOMPTABLE'), memDevise.Taux, memDevise.Code) ;
ChangeMask(FTotalDebit,  memDevise.Decimale, memDevise.Symbole) ;
ChangeMask(FTotalCredit, memDevise.Decimale, memDevise.Symbole) ;
ChangeMask(FTotalSolde,  memDevise.Decimale, memDevise.Symbole) ;
Calcul ;
end;

procedure TZFDevise.BSoldeClick(Sender: TObject);
begin SoldeClick ; end ;

procedure TZFDevise.BSaisTauxClick(Sender: TObject);
begin TauxClick ; end ;

procedure TZFDevise.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
case Key of
      VK_F6 : if (Vide) then begin Key:=0 ; BSoldeClick(nil) ;   end ;
     VK_F10 : if (Vide) then begin Key:=0 ; BValideClick(nil) ;  end ;
     VK_ESCAPE : if (Vide) then begin Key:=0 ; BAbandonClick(nil) ; end ;
  end ;
end;

procedure TZFDevise.FormCreate(Sender: TObject);
begin
InitEntete ;
end;

procedure TZFDevise.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

procedure TZFDevise.FEuroExit(Sender: TObject);
begin
  if (SensDef='D') or (SensDef='M')
    then DEB.SetFocus
    else CRE.SetFocus ;
end;

procedure TZFDevise.BAbandonClick(Sender: TObject);
begin
 Modalresult:=mrCancel ;
end;

procedure TZFDevise.DEBChange(Sender: TObject);
begin
  Calcul ;
end;

procedure TZFDevise.CREChange(Sender: TObject);
begin
  Calcul ;
end;

end.
