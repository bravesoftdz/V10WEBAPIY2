unit RIB;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, DBCGrids, DB, {$IFNDEF DBXPRESS}dbtables,
  HSysMenu, hmsgbox, Hqry, DBCtrls, Hctrls, Grids, DBGrids, Mask{$ELSE}uDbxDataSet{$ENDIF}, Hctrls, DBCtrls,MajTable,
  Mask, hmsgbox, Ent1, HEnt1, Grids, DBGrids, HSysMenu, SaisComm, CpteSav,
  Hqry, UtilPGI ;

Function FicheRIB(Auxiliaire : String ; Mode : TActionFiche ; FromSaisie : boolean ; StRIB : String ; IsAux : Boolean) : integer ;

type
  TFRib = class(TForm)
    TRIB: THTable;
    SRIB  : TDataSource;
    FListe: TDBGrid;
    DBNav : TDBNavigator;
    Panel1    : TPanel;
    FAutoSave : TCheckBox;
    Bevel1    : TBevel;
    BFirst    : TBitBtn;
    BPrev     : TBitBtn;
    BNext     : TBitBtn;
    BLast     : TBitBtn;
    BDelete   : TBitBtn;
    BInsert   : TBitBtn;
    BImprimer : TBitBtn;
    BValider  : TBitBtn;
    BAnnuler  : TBitBtn;
    BAide     : TBitBtn;
    BFerme    : TBitBtn;
    MsgBox    : THMsgBox;
    Pappli    : TPanel;
    TR_PAYS          : THLabel;
    R_PAYS           : THDBValComboBox;
    TR_VILLE         : THLabel;
    R_VILLE          : TDBEdit;
    TR_NUMERORIB     : THLabel;
    R_NUMERORIB      : TDBEdit;
    R_PRINCIPAL      : TDBCheckBox;
    R_DEVISE         : THDBValComboBox;
    TR_DEVISE        : THLabel;
    GbRib            : TGroupBox;
    TR_CLERIB        : THLabel;
    R_CLERIB         : TDBEdit;
    TR_NUMEROCOMPTE  : THLabel;
    R_NUMEROCOMPTE   : TDBEdit;
    TR_GUICHET       : THLabel;
    R_GUICHET        : TDBEdit;
    TR_ETABBQ        : THLabel;
    R_ETABBQ         : TDBEdit;
    TR_CODEBIC       : THLabel;
    R_CODEBIC        : TDBEdit;
    TRIBR_AUXILIAIRE    : TStringField;
    TRIBR_NUMERORIB     : TIntegerField;
    TRIBR_PRINCIPAL     : TStringField;
    TRIBR_ETABBQ        : TStringField;
    TRIBR_GUICHET       : TStringField;
    TRIBR_NUMEROCOMPTE  : TStringField;
    TRIBR_CLERIB        : TStringField;
    TRIBR_DOMICILIATION : TStringField;
    TRIBR_VILLE         : TStringField;
    TRIBR_PAYS          : TStringField;
    TRIBR_DEVISE        : TStringField;
    TRIBR_CODEBIC       : TStringField;
    TRIBR_SOCIETE       : TStringField;
    XX_WHERE: TPanel;
    HMTrad: THSystemMenu;
    TR_DOMICILIATION: THLabel;
    R_DOMICILIATION: TDBEdit;
    BExportRIB: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SRIBDataChange(Sender: TObject; Field: TField);
    procedure SRIBStateChange(Sender: TObject);
    procedure SRIBUpdateData(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BExportRIBClick(Sender: TObject);
  private
    Auxiliaire : String ;
    LastNum    : Longint ;
    Modifier : Boolean ;
    Mode : TActionFiche ;
    IsAux : Boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  EnregOK : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Procedure ChargeEnreg ;
    Function  VerifiRibBic : Boolean ;
    Function  PrincipalUnique : Boolean ;
  public
    FromSaisie : boolean ;
    StRIB      : String ;
    GeneNum    : integer ;
  end;

implementation

{$R *.DFM}

uses PrintDBG ;

Function FicheRIB(Auxiliaire : String ; Mode : TActionFiche ; FromSaisie : boolean ; StRIB : String ; IsAux : Boolean) : integer ;
var FRib: TFRib;
BEGIN
Result:=-1 ;
FRib:=TFRib.Create(Application) ;
 try
  FRib.Auxiliaire:=Auxiliaire ;
  FRib.Mode:=Mode ;
  FRib.FromSaisie:=FromSaisie ;
  FRib.StRIB:=StRIB ;
  FRib.IsAux:=IsAux ;
  if FRib.Auxiliaire<>''then FRib.ShowModal ;
  Result:=FRIB.GeneNum ;
 Finally
  FRib.free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFRib.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if TRIB.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost)   then Exit ;
  mrNo :  if not Bouge(nbCancel) then Exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFRib.EnregOK : boolean ;
Var RibCalcul : String ;
BEGIN
result:=FALSE  ;
if TRIB.state in [dsInsert,dsEdit]=False then Exit ;
if TRIB.state in [dsInsert,dsEdit] then
   BEGIN
   if TRIBR_DOMICILIATION.AsString='' then
      BEGIN
      MsgBox.Execute(2,'','') ; R_DOMICILIATION.SetFocus ; Exit ;
      END ;
   if TRIBR_VILLE.AsString='' then
      BEGIN
      MsgBox.Execute(3,'','') ; R_VILLE.SetFocus ; Exit ;
      END ;
   if Not VerifiRibBic then Exit ;
   if Not PrincipalUnique then
      BEGIN
      MsgBox.Execute(11,'','') ; R_PRINCIPAL.SetFocus ; Exit ;
      END ;
   If VH^.CtrlRIB Then
      BEGIN
      RibCalcul:=VerifRib(TRIBR_ETABBQ.AsString,TRIBR_GUICHET.AsString,TRIBR_NUMEROCOMPTE.AsString) ;
      If RibCalcul<>TRIBR_CLERIB.AsString Then
         BEGIN
         if MsgBox.execute(12,'','')<>mrYes Then BEGIN R_CLERIB.SetFocus ; R_CLERIB.Text:=RibCalcul ; Exit ; END ;
         END ;
      END ;
   END else Exit ;
Result:=TRUE  ;
END ;

Procedure TFRib.ChargeEnreg ;
var Q : TQuery ;
    Lib : String ;
BEGIN
Q:=OpenSQL('Select T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+AUXILIAIRE+'"',True) ;
if not Q.Eof then Lib:=Q.Fields[0].AsString else Lib:='' ;
Ferme(Q) ;
InitCaption(Self,AUXILIAIRE+' '+Lib,'') ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
END ;

Procedure TFRib.NewEnreg ;
BEGIN
InitNew(TRib) ;
TRIBR_AUXILIAIRE.AsString:=Auxiliaire ;
TRIBR_NumeroRIB.AsInteger:=LastNum+1 ;
R_PAYS.Value:='FRA' ;
R_DOMICILIATION.SetFocus ;
END ;

Function TFRib.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         : if MsgBox.execute(1,'','')<>mrYes then Exit ;
   end ;

if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[5]) ;
result:=TRUE ;
if Button=NbInsert then
   BEGIN
   DBNav.BtnClick(NbLast) ; LastNum:=TRibR_NUMERORIB.AsInteger ;
   DBNav.BtnClick(Button) ; NewEnreg ;
   END ;
END ;

procedure TFRib.FormShow(Sender: TObject);
Var Etab,Guichet,NumCompte,Cle,Dom : String ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TRIB.Open ;
if Auxiliaire<>'' then TRIB.SetRange([Auxiliaire,0],[Auxiliaire,99999]) ;
Case Mode Of
     taConsult : FicheReadOnly(Self) ;
     taCreat..taCreatOne : Bouge(nbInsert) ;
     else if (TRIB.EOF) and (TRIB.BOF) then Bouge(nbInsert) ;
   End ;
if FromSaisie then
   BEGIN
   DecodeRIB(Etab,Guichet,NumCompte,Cle,Dom,StRIB) ; NumCompte:=Trim(NumCompte) ;
   if TRIB.State=dsBrowse then
      begin
      //TRIB.Locate('R_NUMEROCOMPTE',NumCompte,[]) ;
      while not TRIB.EOF do if TRIB.FindField('R_NUMEROCOMPTE').AsString=NumCompte then Break else TRIB.Next ;
      end ;
   END ;
GeneNum:=-1 ;
end;

procedure TFRib.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFRib.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFRib.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFRib.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFRib.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFRib.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFRib.BValiderClick(Sender: TObject);
begin
Modifier:=False ;
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if Mode=taCreatOne then Close ;
   END ;
if FromSaisie then BExportRIBClick(Nil) ;
end;

procedure TFRib.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end ;

procedure TFRib.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRib.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFRib.SRIBDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not TRIB.Modified ; BDelete.Enabled:=Not TRIB.Modified ;
if(TRIB.EOF)And(TRIB.BOF)then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TRIB.BOF;
   DnEnable := Enabled and not TRIB.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
//  code de liaison entre field
//   if ((Field.FieldName='T_LIBELLE') and (T_ABREGE.Field.AsString='')) then
//      T_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFRib.SRIBStateChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFRib.SRIBUpdateData(Sender: TObject);
begin
if Modifier then BEGIN Modifier:=False ; if Not OnSauve then Bouge(nbCancel) ;  END ;
end;

Function TFRib.VerifiRibBic : Boolean ;
BEGIN
Result:=False ;
if R_PAYS.Value='FRA' then
  BEGIN
   if TRIBR_ETABBQ.AsString='' then
      BEGIN MsgBox.Execute(6,'','') ; R_ETABBQ.SetFocus ; Exit ; END ;
   if TRIBR_GUICHET.AsString='' then
      BEGIN MsgBox.Execute(7,'','') ; R_GUICHET.SetFocus ; Exit ; END ;
   if TRIBR_NUMEROCOMPTE.AsString='' then
      BEGIN MsgBox.Execute(8,'','') ; R_NUMEROCOMPTE.SetFocus ; Exit ; END ;
   if TRIBR_CLERIB.AsString='' then
      BEGIN MsgBox.Execute(9,'','') ; R_CLERIB.SetFocus ; Exit ; END ;
  END else
  BEGIN
   if TRIBR_CODEBIC.AsString<>''then BEGIN Result:=True ; Exit ; END else
     if ((TRIBR_ETABBQ.AsString='')OR(TRIBR_GUICHET.AsString='')OR
        (TRIBR_NUMEROCOMPTE.AsString='')OR(TRIBR_CLERIB.AsString='')) then
        BEGIN MsgBox.Execute(10,'','') ; R_CODEBIC.SetFocus ; Exit ; END ;
  END ;
// if Not ControleSaisieRib then BEGIN     Exit ; END ;
Result:=True ;
END ;

Function TFRib.PrincipalUnique : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
if Not R_PRINCIPAL.Checked then Exit ;
Q:=OpenSql('Select R_PRINCIPAL from RIB Where R_AUXILIAIRE="'+Auxiliaire+'" And '+
           'R_PRINCIPAL="X" And R_NUMERORIB<>'+IntToStr(TRIBR_NUMERORIB.AsInteger)+'',True) ;
if Not Q.Eof then Result:=False ;
Ferme(Q) ;
END ;

procedure TFRib.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if((ssCtrl in Shift)AND(Key=VK_DELETE)AND(TRIB.Eof)) then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFRib.BImprimerClick(Sender: TObject);
begin
XX_Where.Hint:=' Where R_AUXILIAIRE="'+Auxiliaire+'"' ;
PrintDBGrid(nil,XX_WHERE,Caption,'PRT_RIB') ; end;

procedure TFRib.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFRib.BExportRIBClick(Sender: TObject);
begin
if TRIB.State=dsBrowse then BEGIN GeneNum:=TRIBR_NUMERORIB.AsInteger ; Close ; END ;
end;

end.
