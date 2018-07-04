unit Saistaux1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ExtCtrls, Buttons, hmsgbox, Ent1, HEnt1, Saisutil,
  HSysMenu, HTB97, HFLabel ;

Type tIdentPieceDev = Record
                      NumP : Integer ;
                      JalP,RefP,LibP,NatP : String ;
                      DateP : TDateTime ;
                      DateMin : TDateTime ;
                      DateMax : TDateTime ;
                      End ;

function SaisieNewTaux1 ( Var DEV : RDEVISE ; Var IdentPD : tidentPieceDev ; Comment : TActionFiche ; Mess : String) : boolean ;
function SaisieNewTaux2000 ( Var DEV : RDEVISE ; DateCpt : TDateTime) : boolean ;

type
  TFSaisTaux1 = class(TForm)
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    GBDEV: TGroupBox;
    HDevise: TLabel;
    TDevise: TLabel;
    HQUOTITE: TLabel;
    TQUOTITE: TLabel;
    HTAUXCHANCEL: TLabel;
    HDATECHANCEL: TLabel;
    TDATECHANCEL: TLabel;
    TTAUXCHANCEL: TLabel;
    HCotation: TLabel;
    TCOTATION: TLabel;
    GBTaux: TGroupBox;
    Label2: TLabel;
    NewTaux: THNumEdit;
    HTitre: TLabel;
    Label4: TLabel;
    NewCotation: THNumEdit;
    HTitre1: TLabel;
    GBPiece: TGroupBox;
    HJournal: TLabel;
    E_JOURNAL: TLabel;
    HNaturePiece: TLabel;
    E_NATUREPIECE: TLabel;
    HDateComptable: TLabel;
    E_DATECOMPTABLE: TLabel;
    HNumeroPiece: TLabel;
    E_NUMEROPIECE: TLabel;
    Dock: TDock97;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BChancelOut: TToolbarButton97;
    HRefinterne: TLabel;
    E_REFINTERNE: TLabel;
    HLibelle: TLabel;
    E_LIBELLE: TLabel;
    FlashPb: TFlashingLabel;
    HAlerte: TLabel;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure NewTauxChange(Sender: TObject);
    procedure NewCotationChange(Sender: TObject);
    procedure BChancelOutClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
    procedure AffichInfos;
    Procedure AffichOrNot(TLbl : TLabel; TVal : TLabel; ValTest : String);
  private
  public
    OnCharge : Boolean ;
    FromSaisie : Boolean ;
    DEV : RDEVISE ;
    IdentPD : tidentPieceDev ;
    Comment : TActionFiche ;
    Mess : String ;
  end;

implementation

{$R *.DFM}

Uses CPChancell_TOF;


function SaisieNewTaux1 ( Var DEV : RDEVISE ; Var IdentPD : tidentPieceDev ; Comment : TActionFiche ; Mess : String) : boolean ;
Var X  : TFSaisTaux1 ;
    ii : integer ;
BEGIN
Result:=False ;
X:=TFSaisTaux1.Create(Application) ;
 Try
  X.DEV:=DEV ;
  X.IdentPD.JalP:=IdentPD.JalP ;
  X.IdentPD.DateP:=IdentPD.DateP ;
  X.IdentPD.DateMin:=IdentPD.DateMin ;
  X.IdentPD.DateMax:=IdentPD.DateMax ;
  X.IdentPD.NatP:=IdentPD.NatP ;
  X.IdentPD.RefP:=IdentPD.RefP ;
  X.IdentPD.LibP:=IdentPD.LibP ;
  X.IdentPD.NumP:=IdentPD.NumP ;
  X.Mess:=Mess ;
  X.Comment:=Comment ;
  X.OnCharge:=TRUE ;
  X.FromSaisie:=FALSE ;
  ii:=X.ShowModal ;
  if ((ii=mrOk) and (X.DEV.Taux>0)) then BEGIN DEV:=X.DEV ; Result:=True ; END ;
 Finally
  X.Free ;
 End ;
END ;


function SaisieNewTaux2000 ( Var DEV : RDEVISE ; DateCpt : TDateTime) : boolean ;
Var X  : TFSaisTaux1 ;
    ii : integer ;
BEGIN
Result:=False ;
X:=TFSaisTaux1.Create(Application) ;
 Try
  X.DEV:=DEV ;
  Fillchar(X.IdentPD,SizeOf(X.IdentPD),#0) ;
  X.IdentPD.DateP:=DateCpt ;
  X.Mess:='' ;
  X.Comment:=taModif ;
  X.OnCharge:=TRUE ;
  X.FromSaisie:=TRUE ;
  ii:=X.ShowModal ;
  if ((ii=mrOk) and (X.DEV.Taux>0)) then
    BEGIN
    DEV.Taux:=X.DEV.Taux ; Result:=True ;
    END ;
 Finally
  X.Free ;
 End ;
END ;


procedure TFSaisTaux1.BValiderClick(Sender: TObject);
begin
NextControl(Self) ;
// YMO 20/01/2006 Msg d'alerte si le taux n'existe pas en chancellerie
IF Not ExisteSQL('SELECT H_TAUXREEL FROM CHANCELL'
                 +' WHERE H_DATECOURS>="'+UsDateTime(VH^.Encours.Deb)
                 // YMO 17/05/2006 FQ17432 Plage de dates pour la vérification du taux
                 +'" AND H_DATECOURS<="'+UsDateTime(VH^.Encours.Fin)
                 +'" AND H_DEVISE ="'+Dev.Code
                 +'" AND H_TAUXCLOTURE ='+VarianttoSQL(StrToFloat(NewTaux.Text))) then
begin
  If HM.Execute(12,Caption,'')<>mrYes then begin NewTaux.SetFocus; Exit; end;
end
else
  If (Not FromSaisie) And (HM.Execute(6,Caption,'')<>mrYes) Then Exit ;

if NewTaux.Value<=0 then BEGIN HM.Execute(0,Caption,'') ; Exit ; END ;
if NewCotation.Value<=0 then BEGIN HM.Execute(0,Caption,'') ; Exit ; END ;
//if (ActiveControl=NewTaux) Or  (ActiveControl=NewCotation) then BValider.SetFocus ;
DEV.Taux:=NewTaux.Value*V_PGI.TauxEuro ;
Dev.Cotation:=1/NewCotation.Value ;
ModalResult:=mrOk ;
end;

procedure TFSaisTaux1.FormShow(Sender: TObject);
begin
OnCharge:=TRUE ;
Caption:=HM.Mess[8] ;
If FromSaisie Then
  BEGIN
  GBPiece.Visible:=FALSE ;
  Height:=Height-GBPiece.Height ;
  Caption:=HM.Mess[9] ;
  END ;
If Comment=taConsult Then
  BEGIN
  BValider.Visible:=FALSE ; NewCotation.Enabled:=FALSE ; NewTaux.Enabled:=FALSE ; Caption:=HM.Mess[7] ;
  END ;
If (IdentPD.DateP<>iDate1900) {si plusieurs écritures} And (IdentPD.DateP<V_PGI.DateDebutEuro) Then
  BEGIN
  HM.Execute(5,Caption,'') ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;

NewTaux.Value:=Arrondi(DEV.Taux/V_PGI.TauxEuro,9) ;
If Arrondi(Dev.Cotation,5)=0 Then
BEGIN
  If Arrondi(Dev.Taux,5)=0 Then NewCotation.Value:=1  Else NewCotation.Value:=Arrondi(1/NewTaux.Value,9) ;
END
ELSE
  NewCotation.Value:=Arrondi(1/DEV.Cotation,9) ;

AffichInfos;

HTitre.Caption:=HM.Mess[1]+' '+TDevise.Caption+' '+HM.Mess[2] ;
HTitre.Visible:=True ;
HTitre1.Caption:=HM.Mess[3]+' '+TDevise.Caption+' '+HM.Mess[4] ;
HTitre1.Visible:=True ;

HAlerte.Caption:=HM.Mess[11];

OnCharge:=FALSE ;
If Mess<>'' Then
  BEGIN
  FlashPb.Visible:=True ; FlashPb.Flashing:=True ; FlashPb.Caption:=Mess ;
  END ;
UpdateCaption(Self) ;
end;

procedure TFSaisTaux1.NewTauxChange(Sender: TObject);
begin
If OnCharge Then Exit ;
If (Arrondi(NewTaux.Value,9)<>0) Then NewCotation.Value:=Arrondi(1/NewTaux.Value,9) Else NewCotation.Value:=0 ;
end;

procedure TFSaisTaux1.NewCotationChange(Sender: TObject);
begin
If OnCharge Then Exit ;
If (Arrondi(NewCotation.Value,9)<>0) Then NewTaux.Value:=Arrondi(1/NewCotation.Value,9) Else NewTaux.Value:=0 ;
end;


procedure TFSaisTaux1.BChancelOutClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
{$ENDIF} // YMO 05/01/06 Modifications des ecritures : pquoi avait-on squizzé l'appel en eAGL ?
FicheChancel(Dev.Code,TRUE,0,taModif,True) ;

end;

procedure TFSaisTaux1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
Case Key of
   VK_ESCAPE : BEGIN Close ;  END ;
   VK_F10    : BEGIN BValiderClick(Nil) ; Key:=0 ; END ;
   END ;

end;

procedure TFSaisTaux1.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end ;

procedure TFSaisTaux1.AffichInfos;
begin

TDEVISE.Caption:=RechDom('ttDevisetoutes',DEV.Code,False) ;

TQUOTITE.Caption:=FloatToStr(DEV.Quotite) ; // non renseigné à ce stade
//AffichOrNot(HQUOTITE, TQUOTITE, FloatToStr(DEV.Quotite));

AffichOrNot(HDATECHANCEL, TDATECHANCEL, DateToStr(DEV.DateTaux));
AffichOrNot(HTAUXCHANCEL, TTAUXCHANCEL, FloatToStr(NewTaux.Value));
AffichOrNot(HCOTATION, TCOTATION, FloatToStr(NewCotation.Value));
AffichOrNot(HJOURNAL, E_JOURNAL, IdentPD.JalP);
AffichOrNot(HNUMEROPIECE, E_NUMEROPIECE, IntToStr(IdentPD.NumP));
AffichOrNot(HDATECOMPTABLE, E_DATECOMPTABLE, DateToStr(IdentPD.DateP));
AffichOrNot(HREFINTERNE, E_REFINTERNE, IdentPD.RefP);
AffichOrNot(HLIBELLE, E_LIBELLE, IdentPD.LibP);
AffichOrNot(HNATUREPIECE, E_NATUREPIECE, IdentPD.NatP);

HAlerte.Visible:=DEV.DateTaux=iDate1900;

end ;


Procedure TFSaisTaux1.AffichOrNot(TLbl : TLabel; TVal : TLabel; ValTest : String);
begin
  TVal.Caption := ValTest;
  If (ValTest = '') or (ValTest = '0') or (ValTest=StDate1900) then
  begin
      TVal.Visible := False;
      TLbl.Visible := False;
  end
  else
  begin
      TVal.Visible := True;
      TLbl.Visible := True;
  end;
end;

end.
