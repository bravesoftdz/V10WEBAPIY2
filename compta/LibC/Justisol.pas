unit Justisol;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Hctrls, Mask, Hcompte, StdCtrls, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, Ent1, HEnt1, SaisUtil,
  SaisComm,
{$IFNDEF IMP}
  Saisie,
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HRichOLE,UtilPGI,
  ADODB, HPanel, hmsgbox;

procedure JustifSolde(LeCpte : String ; Lefb : TFichierBase );

type
  TFJustisol =  class(TFMul)
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_ETATLETTRAGE: THLabel;
    E_ETATLETTRAGE: THValComboBox;
    E_GENERAL: THCpteEdit;
    E_AUXILIAIRE: THCpteEdit;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_DATEECHEANCE_: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE: THLabel;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    PPied: TPanel;
    TT_LIBELLE: TLabel;
    JU_SOLDEG: THNumEdit;
    JU_TOTALDEBIT: THNumEdit;
    JU_TOTALCREDIT: THNumEdit;
    E_QUALIFPIECE: THValComboBox;
    TE_QUALIFPIECE: THLabel;
    E_REFINTERNE: TEdit;
    TE_REFINTERNE: THLabel;
    MsgBox: THMsgBox;
    procedure BChercheClick(Sender: TObject); override;
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BNouvRechClick(Sender: TObject);
  private    { Déclarations privées }
    LeCpte : String ;
    Lefb : TFichierbase ;
    Procedure LibelleAuxi ;
    Procedure CalculLeSolde ;
    procedure InitCriteres ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;

procedure JustifSolde(LeCpte : String ; Lefb : TFichierbase);
var FJustisol : TFJustisol ;
begin
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FJustisol:=TFJustisol.Create(Application) ;
 try
  FJustisol.LeCpte:=LeCpte ;
  FJustisol.Lefb:=Lefb ;
  FJustisol.FNomFiltre:='JUTIFSLDX' ;
  FJustisol.Q.Manuel:=TRUE ;
  FJustisol.Q.Liste:='CPJUSTIFSLDA' ;
  FJustisol.ShowModal ;
 finally
  FJustisol.Free ;
 end;
Screen.Cursor:=crDefault ;
end;

procedure TFJustisol.BChercheClick(Sender: TObject);
begin
  inherited;
//Q.UpdateCriteres ; CentreDBGrid(FListe) ;
LibelleAuxi ; FListe.Visible:=False ; CalculLeSolde ; FListe.Visible:=True ;
end;

procedure TFJustisol.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFJustisol.FormShow(Sender: TObject);
Var D1,D2 : TDateTime ;
begin
E_DATECOMPTABLE.Text:=StDate1900 ; E_DATECOMPTABLE_.Text:=StDate2099 ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
InitCriteres ;
Case Lefb of
   fbAux  :BEGIN
           E_AUXILIAIRE.Text:=LeCpte ; E_GENERAL.Text:='' ;
           Caption:=MsgBox.Mess[1]+LeCpte ;
           END ;
   fbGene :BEGIN
           E_AUXILIAIRE.Text:='' ; E_GENERAL.Text:=LeCpte ; Q.Liste:='CPJUSTIFSLDG' ;
           Caption:=MsgBox.Mess[0]+LeCpte ;
           END ;
   End ;
E_EXERCICE.Value:=QuelDateExo(V_PGI.DateEntree,D1,D2) ;
ExoToDates (E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
E_DEVISE.Value:=V_PGI.DevisePivot ;
E_ETABLISSEMENT.ItemIndex:=0 ; E_ETATLETTRAGE.ItemIndex:=0 ; E_QUALIFPIECE.ItemIndex:=0 ;
  inherited;
Q.Manuel:=FALSE ; ChercheClick ;
LibelleAuxi ; FListe.Visible:=False ; CalculLeSolde ; FListe.Visible:=True ;
ChangeMask(JU_TOTALDEBIT,V_PGI.OkDecV,V_PGI.SymbolePivot) ;
ChangeMask(JU_TOTALCREDIT,V_PGI.OkDecV,V_PGI.SymbolePivot) ;
UpdateCaption(Self) ;
end;

Procedure TFJustisol.LibelleAuxi ;
Var Q1 : TQuery ;
    St : String ;
BEGIN
st:='' ;
if Q.Eof then BEGIN TT_LIBELLE.Caption:='' ; Exit ; END ;
Case lefb of
    fbAux  :BEGIN
            Q1:=OpenSql('Select T_LIBELLE from TIERS where T_AUXILIAIRE="'+LeCpte+'"',True) ;
            St:=Q1.Fields[0].AsString ; Ferme(Q1) ;
            END ;
    fbGene :BEGIN
            Q1:=OpenSql('Select G_LIBELLE from GENERAUX where G_GENERAL="'+LeCpte+'"',True) ;
            st:=Q1.Fields[0].AsString ; Ferme(Q1) ;
            END ;
    End ;
TT_LIBELLE.Caption:=St ;
END ;

Procedure TFJustisol.CalculLeSolde ;
Var Sc,Sd : Extended ;
BEGIN
Q.First ; Sc:=0 ; Sd:=0 ;
While Not Q.EOF do
  BEGIN
  Sc:=Sc+Q.FindField('E_CREDIT').AsFloat ;
  Sd:=Sd+Q.FindField('E_DEBIT').AsFloat ;
  Q.Next ;
  END ;
JU_TOTALDEBIT.Value:=Sd ; JU_TOTALCREDIT.Value:=Sc ;
AfficheLeSolde(JU_SOLDEG,Sd,Sc) ; Q.First ;
END ;

procedure TFJustisol.FListeDblClick(Sender: TObject);
begin
  inherited;
{$IFNDEF IMP}
TrouveEtLanceSaisie(Q,taConsult,'') ;
{$ENDIF}
end;

procedure TFJustiSol.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
Case Lefb of
   fbAux  :BEGIN
           E_AUXILIAIRE.Text:=LeCpte ; E_GENERAL.Text:='' ;
           END ;
   fbGene :BEGIN
           E_AUXILIAIRE.Text:='' ; E_GENERAL.Text:=LeCpte ; Q.Liste:='CPJUSTIFSLDG' ;
           END ;
   End ;
END ;

procedure TFJustisol.BNouvRechClick(Sender: TObject);
begin
  inherited;
InitCriteres ;
end;

end.
