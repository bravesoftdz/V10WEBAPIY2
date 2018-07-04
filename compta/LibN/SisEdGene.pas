unit SisEdGene;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QRS1, Db, HMemTb, ExtCtrls, DBTables, Hqry, Menus, HSysMenu, hmsgbox,
  Grids, DBGrids, HDB, HQuickrp, StdCtrls, Hctrls, HTB97, Buttons, ComCtrls,
  HPanel,UiUtil,Hent1, Mask, Hcompte,ParamDat,Ent1,EdtEtat,LookUp,
  HPdfviewer;

type
  TFSisEdGene = class(TFQRS1)
    E_JOURNAL: THCritMaskEdit;
    TE_JOURNAL: THLabel;
    TE_JOURNAL_: THLabel;
    E_JOURNAL_: THCritMaskEdit;
    E_EXERCICE: THValComboBox;
    HLabel1: THLabel;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    HLabel2: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    HLabel3: THLabel;
    E_DEVISE: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    TE_NUMEROPIECE2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    function RecupWhereSQL : string ;override;
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure FormDestroy(Sender: TObject);
  private
    { Déclarations privées }
    fTypeEdition : string;
    TEXO2        : THLabel ;
    EXO2         : THValComboBox ;
    TE_GENERAL   : THLabel ;
    E_GENERAL    : THCritMaskEdit ;
    TE_GENERAL2  : THLabel ;
    E_GENERAL_   : THCritMaskEdit ;
    procedure FillComboBds ;
    procedure AjoutCriteres ;
    procedure DetruitCriteres ;
  public
    { Déclarations publiques }
  end;

procedure LanceEditionSISS57 (TypeEdition : string);

implementation

{$R *.DFM}

procedure LanceEditionSISS57 (TypeEdition : string);
var X : TFSisEdGene ; PP : THPanel; Q : TQuery ;
begin
  if TypeEdition='BM1' then
    begin
    Q:=OpenSQL('SELECT DISTINCT HB_EXERCICE FROM HISTOBAL WHERE HB_TYPEBAL="BDS"', TRUE) ;
    if Q.EOF then
      begin
      Ferme(Q) ;
      HShowMessage('5;Erreur;Aucune balance de situation;E;O;O;O;', '', '') ;
      Exit ;
      end ;
    Ferme(Q) ;
    end ;
  X := TFSisEdGene.Create (Application);
  PP:=FindInsidePanel ;
  X.fTypeEdition := TypeEdition;
  X.InitQR('E','UCO',TypeEdition, True,True) ;
  if PP=Nil then
  begin
    try
      X.ShowModal ;
     finally
      X.Free ;
     end ;
   END else
   BEGIN
     InitInside(X,PP) ;
     X.Show ;
   end;
   Screen.Cursor:=SyncrDefault ;
end;

procedure TFSisEdGene.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if isInside(Self) then Action:=caFree ;
end;

procedure TFSisEdGene.FormShow(Sender: TObject);
begin
inherited ;
bEuro.Visible:=FALSE ;
UpdateCaption(self);
E_EXERCICE.Value:=VH^.Entree.Code ;
E_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
E_DEVISE.Value:=V_PGI.DevisePivot ;
if fTypeEdition='BM1' then AjoutCriteres ;
end;

function TFSisEdGene.RecupWhereSQL : string ;
begin
Result:=inherited RecupWhereSQL ;
end;

procedure TFSisEdGene.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFSisEdGene.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
ParamDate(Self,Sender,Key) ;
end;

procedure TFSisEdGene.FillComboBds ;
var Q : TQuery ; Cols, Where : string ;
begin
EXO2.Clear ; EXO2.Values.Clear ;
Cols:='HB_EXERCICE, HB_DATE1, HB_DATE2' ;
Where:='HB_TYPEBAL="BDS"' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM HISTOBAL WHERE '+Where, TRUE) ;
while not Q.EOF do
  begin
  EXO2.Values.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';'+Q.Fields[2].AsString+';') ;
  EXO2.Items.Add(TraduireMemoire('Balance du '+DateToStr(Q.Fields[1].AsDateTime)+' au '+DateToStr(FinDeMois(Q.Fields[2].AsDateTime)))) ;
  Q.Next ;
  end ;
Ferme(Q) ;
EXO2.ItemIndex:=0 ;
end ;

procedure TFSisEdGene.AjoutCriteres ;
begin
// TPageControl "plus grand"
Pages.Height:=Pages.Height+30 ;
// Balances comparatives
TEXO2:=THLabel.Create(Self) ;              TEXO2.Parent:=Standards ;
TEXO2.Name:='LCOMPARE' ;                   TEXO2.Caption:=TraduireMemoire('Com&paratif') ;
TEXO2.Top:=TE_DATECOMPTABLE.Top+25 ;       TEXO2.Left:=TE_DATECOMPTABLE.Left ;
TEXO2.Width:=58 ;                          TEXO2.Visible:=TRUE ;
EXO2:=THValComboBox.Create(Self) ;         EXO2.Parent:=Standards ;
EXO2.Name:='EXO2' ;                        //EXO2.DataType:='TTEXERCICE' ;
EXO2.Style:=csDropDownList ;               EXO2.Vide:=FALSE ;
EXO2.Top:=E_DATECOMPTABLE.Top+25 ;         EXO2.Left:=E_DATECOMPTABLE.Left ;
EXO2.Width:=E_EXERCICE.Width ;             EXO2.Visible:=TRUE ;
FillComboBds ;
// Comptes
TE_GENERAL:=THLabel.Create(Self) ;         TE_GENERAL.Parent:=Standards ;
TE_GENERAL.Name:='TE_GENERAL' ;            TE_GENERAL.Caption:=TraduireMemoire('&Compte de') ;
TE_GENERAL.Top:=TE_NUMEROPIECE.Top+25 ;    TE_GENERAL.Left:=TE_NUMEROPIECE.Left ;
TE_GENERAL.Width:=58 ;                     TE_GENERAL.Visible:=TRUE ;
TE_GENERAL2:=THLabel.Create(Self) ;        TE_GENERAL2.Parent:=Standards ;
TE_GENERAL2.Name:='TE_GENERAL2' ;          TE_GENERAL2.Caption:=TraduireMemoire('à') ;
TE_GENERAL2.Top:=TE_NUMEROPIECE2.Top+25 ;  TE_GENERAL2.Left:=TE_NUMEROPIECE2.Left ;
TE_GENERAL2.Width:=58 ;                    TE_GENERAL2.Visible:=TRUE ;
E_GENERAL:=THCritMaskEdit.Create(Self) ;   E_GENERAL.Parent:=Standards ;
E_GENERAL.Name:='E_GENERAL' ;              E_GENERAL.DataType:='TZGENERAL' ;
E_GENERAL.Text:='' ;
E_GENERAL.OpeType:=otString ;              E_GENERAL.Operateur:=Superieur ;
E_GENERAL.Top:=E_NUMEROPIECE.Top+25 ;      E_GENERAL.Left:=E_NUMEROPIECE.Left ;
E_GENERAL.Width:=E_NUMEROPIECE.Width ;     E_GENERAL.Visible:=TRUE ;
E_GENERAL.ElipsisButton:=TRUE ;            E_GENERAL.ElipsisAutoHide:=TRUE ;
E_GENERAL_:=THCritMaskEdit.Create(Self) ;  E_GENERAL_.Parent:=Standards ;
E_GENERAL_.Name:='E_GENERAL_' ;            E_GENERAL_.DataType:='TZGENERAL' ;
E_GENERAL_.Text:='' ;
E_GENERAL_.OpeType:=otString ;             E_GENERAL_.Operateur:=Inferieur ;
E_GENERAL_.ElipsisButton:=TRUE ;           E_GENERAL_.ElipsisAutoHide:=TRUE ;
E_GENERAL_.Top:=E_NUMEROPIECE_.Top+25 ;    E_GENERAL_.Left:=E_NUMEROPIECE_.Left ;
E_GENERAL_.Width:=E_NUMEROPIECE_.Width ;   E_GENERAL_.Visible:=TRUE ;
end ;

procedure TFSisEdGene.DetruitCriteres ;
begin
TEXO2.Free ;
EXO2.Free ;
TE_GENERAL.Free ;
E_GENERAL.Free ;
TE_GENERAL2.Free ;
E_GENERAL_.Free ;
end ;

procedure TFSisEdGene.FormDestroy(Sender: TObject);
begin
inherited ;
if fTypeEdition='BM1' then DetruitCriteres ;
end;

end.

