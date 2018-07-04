unit TvaAcc ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, Ent1,
  HEnt1, Saisie, Saisutil, SaisBase, hmsgbox, TvaLettr, SaisComm,CritEdt,UtilEdt,
  HTB97, ColMemo, HPanel, UiUtil, HRichOLE, ADODB ;

Procedure TvaModifAcc ;

type
  TFTvaAcc = class(TFMul)
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_DATECREATION: THLabel;
    E_DATECREATION: THCritMaskEdit;
    TE_DATECREATION_: THLabel;
    E_DATECREATION_: THCritMaskEdit;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_MODEPAIE: THLabel;
    E_MODEPAIE: THValComboBox;
    TE_REGIMETVA: THLabel;
    E_REGIMETVA: THValComboBox;
    TE_LIBELLE: THLabel;
    E_LIBELLE: TEdit;
    E_NUMECHE: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    E_ECRANOUVEAU: THCritMaskEdit;
    XX_WHERE: TEdit;
    E_ECHE: THCritMaskEdit;
    BZoomPiece: TToolbarButton97;
    HM: THMsgBox;
    E_EDITEETATTVA: THCritMaskEdit;
    XX_WHERE1: TEdit;
    RAcc: TRadioGroup;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    XX_WHERE2: TEdit;
    BDetailFact: TToolbarButton97;
    XX_WHERE3: TEdit;
    POPZ: TPopupMenu;
    BImpListe: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure RAccClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BDetailFactClick(Sender: TObject);
    procedure BImpListeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
  public
  end;

implementation

uses (*QRTvaAcc,*) UtilPgi ;

{$R *.DFM}

Procedure TvaModifAcc ;
Var X  : TFTvaAcc ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
X:=TFTvaAcc.Create(Application) ;
X.FNomFiltre:='TVAACC' ; X.Q.Manuel:=TRUE ; X.Q.Liste:='TVAACC' ;
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



procedure TFTvaAcc.FormShow(Sender: TObject);
begin
E_DEVISE.Value:=V_PGI.DevisePivot ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
E_DATECREATION.Text:=StDate1900 ; E_DATECREATION_.Text:=StDate2099 ;
PositionneEtabUser(E_ETABLISSEMENT) ;
if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then
   BEGIN
   if Not E_JOURNAL.Vide then E_JOURNAL.Value:=E_JOURNAL.Values[0] else
    if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
   END ;
XX_WHERE3.Text:='E_DEVISE="'+V_PGI.DevisePivot+'"' ;    
  inherited;
end;

procedure TFTvaAcc.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFTvaAcc.BZoomPieceClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Text) ;
end;

procedure TFTvaAcc.RAccClick(Sender: TObject);
begin
  inherited;     {17/09/2007 YMO Ajout des règlements à la requête}
if Racc.ItemIndex=0 then XX_WHERE2.Text:='E_NATUREPIECE="OC" OR E_NATUREPIECE="RC"'
                    else XX_WHERE2.Text:='E_NATUREPIECE="OF" OR E_NATUREPIECE="RF"' ;
end;

procedure TFTvaAcc.FListeDblClick(Sender: TObject);
begin
  inherited;
BZoomPieceClick(Nil) ;
end;

procedure TFTvaAcc.BDetailFactClick(Sender: TObject);
Var M : RMVT ;
    QEcr : TQuery ;
    OEcr : TOBM ;
begin
  inherited;
if Not TrouveSaisie(Q,M,E_QUALIFPIECE.Text) then Exit ;
OEcr:=Nil ;
M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
M.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
QEcr:=OpenSQL('Select * from ECRITURE Where '+WhereEcriture(tsGene,M,True),True) ;
if Not QEcr.EOF then
   BEGIN
   OEcr:=TOBM.Create(EcrGen,'',True) ;
   OEcr.ChargeMvt(QEcr) ;
   END ;
Ferme(QEcr) ;
if OEcr<>Nil then
   BEGIN
   TvaAccLettrage(OEcr) ;
   OEcr.Free ;
   END ;
end;

procedure TFTvaAcc.BImpListeClick(Sender: TObject);
var Crit : TCritEdt ;
    D1,D2 : TDateTime ;
begin
  inherited;
Fillchar(Crit,SizeOf(Crit),#0) ;
Crit.NatureEtat:=neGL ;
InitCritEdt(Crit) ;
D1:=StrToDate(E_DATECOMPTABLE.Text) ; D2:=StrToDate(E_DATECOMPTABLE_.Text) ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.GL.Deductible:=RAcc.ItemIndex=1 ;
//EditionTvaAcc(Crit) ;
end;

procedure TFTvaAcc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then _Bloqueur('nrSaisieModif',False) ;
  inherited;
end;

end.
