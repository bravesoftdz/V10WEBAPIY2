{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/04/2003
Modifié le ... :   /  /    
Description .. : Remplace en eAGL par CPCHANCELL_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit CHANCEL;

interface                          

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, DBCtrls, StdCtrls, Buttons, ExtCtrls, Spin, HDB, Mask,
  Hctrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, Grids, DBGrids, Ent1, hmsgbox, Hqry, Purge,
  hcompte, HEnt1, HSysMenu, MajTable, SaisUtil, ADODB, TntDBGrids,
  TntButtons, TntStdCtrls ;

Procedure FicheChancel(Quel : String ; JusteUne : Boolean ; DateDeb : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;
Procedure FicheChancelsur2dates(Quel : String ; DateDeb,DateFin : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;

type
  TFChancel = class(TForm)
    SChancell: TDataSource;
    TChancell: THTable;
    PDevise: TPanel;
    MsgBox: THMsgBox;
    panel: TPanel;
    BInsert: THBitBtn;
    HelpBtn: THBitBtn;
    BPurge: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    FAutoSave: TCheckBox;
    BFerme: THBitBtn;
    BDelete: THBitBtn;
    BImprimer: THBitBtn;
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    FListe: THDBGrid;
    DBNav: TDBNavigator;
    TChancellH_DATECOURS   : TDateTimeField;
    TChancellH_TAUXREEL    : TFloatField;
    TChancellH_TAUXCLOTURE : TFloatField;
    TChancellH_COMMENTAIRE : TStringField;
    TChancellH_DEVISE      : TStringField;
    HMTrad: THSystemMenu;
    TH_DATEDU: THLabel;
    H_DATEDU: TMaskEdit;
    TH_DATEAU: THLabel;
    H_DATEAU: TMaskEdit;
    TQuotite: THLabel;
    TValeurQ: THLabel;
    OkDate: TCheckBox;
    CodeDevise: THValComboBox;
    TH_DEVISE: THLabel;
    TValable: THLabel;
    RGSens: TRadioGroup;
    TChancellH_COTATION: TFloatField;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure CodeDeviseChange(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BPurgeClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure SChancellDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TChancellH_TAUXREELChange(Sender: TField);
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure OkDateClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TChancellNewRecord(DataSet: TDataSet);
    procedure TChancellPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure FListeRowEnter(Sender: TObject);
    procedure SChancellUpdateData(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure H_DATEDUExit(Sender: TObject);
    procedure H_DATEAUExit(Sender: TObject);
    procedure RGSensClick(Sender: TObject);
    procedure TChancellBeforePost(DataSet: TDataSet);
  private
    Lequel : String ;
    JustOne : Boolean ;
    FirstDate : TDateTime ;
    LastDate  : TDateTime ;
    Action    : TActionFiche ;
    Sur2Dates : Boolean ;
    SurEuro,MonnaieIn : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure ChargeEnreg ;
    Function  QuotiteChange : Integer ;
    procedure PurgeDates( DatesSup : TDateTime ) ;
//    Function  DateRecente( DateOk : TDateTime ) : Boolean ;
    Function  ChercheEnreg(QuelDate : TDateTime) : Boolean ;
    Function  LimiteDateEuro(QuelDate : TDateTime) : boolean ;
    procedure ChangeIndex;
    procedure InitSur2Dates ;
    procedure GereSens ( ForceCol : boolean ) ;
  public
  end;

implementation

uses PrintDBG,Devise,UtilPGI;

{$R *.DFM}

procedure TFChancel.InitSur2Dates ;
BEGIN
If Sur2Dates And SurEuro Then
  BEGIN
  If FirstDate<V_PGI.DateDebutEuro Then FirstDate:=V_PGI.DateDebutEuro ;
  If LastDate<V_PGI.DateDebutEuro Then LastDate:=V_PGI.DateDebutEuro+1 ;
  If Double(FirstDate)>0 Then H_DATEDU.Text:=DateToStr(FirstDate) ;
  If Double(LastDate)>0 Then H_DATEAU.Text:=DateToStr(LastDate) ;
  OkDate.Checked:=TRUE ;
  END ;
END ;

procedure TFChancel.FormShow(Sender: TObject);
Var QLoc : TQuery ;
    DD : TDateTime ;
begin
// Appel de la fonction d'empilage dans la liste des fiches
AglEmpileFiche(Self) ;
H_DateDu.Text:=StDate1900 ; H_DateAu.Text:=StDate2099 ;
if Not SurEuro then
   BEGIN
   H_DATEAU.Text:=DateToStr(V_PGI.DatedebutEuro-1) ;
   RGSens.ItemIndex:=1 ; RGSens.Visible:=False ;
   END ;
GereSens(False) ;
if SurEuro then HelpContext:=1150200 ;
CodeDevise.Values.Clear ; CodeDevise.Items.Clear ;
QLoc:=OpenSql('Select D_DEVISE,D_LIBELLE From DEVISE Where D_DEVISE<>"'+V_PGI.DevisePivot+'" Order by D_DEVISE',True,-1,'',true);
While Not QLoc.Eof do
   BEGIN
   CodeDevise.Values.Add(QLoc.Fields[0].AsString) ; CodeDevise.Items.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
TChancell.Open ;
if Lequel<>'' then
   BEGIN
   if (Action<>tacreat) And (Not Sur2Dates) then
      BEGIN
      if FirstDate>0 then DD:=FirstDate else DD:=V_PGI.DateEntree ;
      if not FindLaKey(TChancell,[Lequel,DD]) then
         BEGIN
         if ChercheEnreg(DD) then
            BEGIN
            MessageAlerte(MsgBox.Mess[7]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
            END else Action:=taCreat ;
         END ;
      END ;
   CodeDevise.Value:=Lequel ;
   If JustOne Then CodeDevise.Enabled:=FALSE ;
   If (Double(FirstDate)>0) Then
      BEGIN
      H_DATEDU.Text:=DateToStr(FirstDate) ;
      OkDate.Checked:=TRUE ;
      END else
      BEGIN
      if MonnaieIn then
        BEGIN
        H_DATEAU.Text:=DateToStr(V_PGI.DateDebutEuro-1) ;
        END else
        BEGIN
        if SurEuro then H_DATEDU.Text:=DateToStr(V_PGI.DateDebutEuro)
                   else H_DATEAU.Text:=DateToStr(V_PGI.DateDebutEuro-1) ;
        END ;
      END ;
   InitSur2Dates ;
   ChangeIndex ;
   if Action=taCreat Then Bouge(nbInsert) ;
   END else
   BEGIN
   If CodeDevise.Values.Count>0 Then CodeDevise.Value:=CodeDevise.Values[0] ;
   If Sur2Dates And SurEuro Then BEGIN InitSur2Dates ; ChangeIndex ; END ;
   END ;
if Action=taConsult then BEGIN FicheReadOnly(Self) ; BPurge.Enabled:=False ; END ;
end;


Procedure FicheChancel(Quel : String ; JusteUne : Boolean ; DateDeb : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;
var FChancel : TFChancel ;
begin
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FChancel:= TFChancel.Create(Application) ;
try
  FChancel.Lequel:=Quel ; FCHancel.JustOne:=JusteUne ;
  FChancel.FirstDate:=0 ; FChancel.LastDate:=0 ;
  FChancel.Action:=Action ; FChancel.Sur2Dates:=FALSE ;
  FChancel.SurEuro:=SurEuro ;
  If Double(DateDeb)>0 Then BEGIN FChancel.FirstDate:=DateDeb ; END ;
  FChancel.ShowModal ;
  finally
  FChancel.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

Procedure FicheChancelSur2Dates(Quel : String ; DateDeb,DateFin : TDateTime ; Action : TActionFiche ; SurEuro : boolean) ;
var FChancel : TFChancel ;
begin
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FChancel:= TFChancel.Create(Application) ;
try
  FChancel.Lequel:=Quel ; FCHancel.JustOne:=FALSE ;
  FChancel.FirstDate:=0 ; FChancel.Action:=Action ;
  FChancel.SurEuro:=SurEuro ; FChancel.Sur2Dates:=TRUE ;
  If Double(DateDeb)>0 Then BEGIN FChancel.FirstDate:=DateDeb ; END ;
  If Double(DateFin)>0 Then BEGIN FChancel.LastDate:=DateFin ; END ;
  FChancel.ShowModal ;
  finally
  FChancel.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFChancel.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFChancel.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFChancel.ChangeIndex;
Var TypeDev : string ;
    DateDeb,DateFin : TDateTime ;
begin
if Not TChancell.Active then exit ;
if IsvalidDate(H_DATEDU.Text) then
   if IsvalidDate(H_DATEAU.Text) then
      BEGIN
      TypeDev:=CodeDevise.Value ;
      If OkDate.Checked Then
         BEGIN
         DateDeb:=StrToDate(H_DATEDU.Text) ;
         DateFin:=StrToDate(H_DATEAU.Text) ;
         END Else
         BEGIN
         DateDeb:=IDate1900 ;
         DateFin:=IDate2099 ;
         if SurEuro then DateDeb:=V_PGI.DateDebutEuro else DateFin:=V_PGI.DateDebutEuro-1 ;
         END ;
      TChancell.SetRange([TypeDev,DateDeb],[TypeDev,DateFin]) ;
      END ;
end ;

procedure TFChancel.CodeDeviseChange(Sender: TObject);
var Libelle0,Libelle1 : string;
begin
  Libelle0 := MsgBox.Mess[16];
  Libelle1 := MsgBox.Mess[17];
  //
  ChangeIndex ;
  MonnaieIn:=EstMonnaieIn(CodeDevise.Value) ;
  if MonnaieIN then
  BEGIN
	  RGSens.ItemIndex:=1 ; RGSens.Visible:=False ; RGSensClick(Nil) ;
  END else
  BEGIN
    RGSens.Visible:=True ;
  END ;
  if CodeDevise.Value<>'' then
  BEGIN
    Libelle0:=FindEtReplace(Libelle0,'<$$$>',CodeDevise.Value,True) ;
    Libelle1:=FindEtReplace(Libelle1,'<$$$>',CodeDevise.Value,True) ;
  END ;
  if V_PGI.DevisePivot <>'' then
  BEGIN
    Libelle0:=FindEtReplace(Libelle0,'<FFF>',V_PGI.DevisePivot,true) ;
    Libelle1:=FindEtReplace(Libelle1,'<FFF>',V_PGI.DevisePivot,True) ;
  END ;
  RGsens.Items[0]:=libelle0;
  RGsens.Items[1]:=Libelle1 ;
end;

procedure TFChancel.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFChancel.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFChancel.BPurgeClick(Sender: TObject);
Var LesDatesASup : TDateTime ;
    Depuis : String ;
begin
if TChancellH_DATECOURS.AsString<>'' then
   BEGIN
   Depuis:=TChancellH_DATECOURS.AsString ;
   if PurgeOui(Depuis,CodeDevise.Value,LesDatesASup) then
      BEGIN
      Purgedates(LesDatesASup) ;
      TChancell.Close ; TChancell.Open ;
      TChancell.SetRange([CodeDevise.Value,IDate1900],[CodeDevise.Value,IDate2099]) ;
      END ;
   END ;
end;

procedure TFChancel.PurgeDates(DatesSup : TDateTime) ;
BEGIN
ExecuteSQl('DELETE FROM CHANCELL WHERE H_DEVISE="'+TChancellH_DEVISE.AsString+'" '+
           ' AND H_DATECOURS<="'+USDatetime(DatesSup)+'" ') ;
END ;

procedure TFChancel.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFChancel.SChancellDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=(Not(TChancell.State in [dsEdit,dsInsert])) ;
BDelete.Enabled:=(Not(TChancell.State in [dsEdit,dsInsert])) ;
BPurge.Enabled:=(Not(TChancell.State in [dsEdit,dsInsert])) ;
if(TChancell.Eof)And(TChancell.Bof) then BDelete.Enabled:=False ;
CodeDevise.Enabled:=(Not(TChancell.State in [dsEdit,dsInsert])) ;
OkDate.Enabled:=(Not(TChancell.State in [dsEdit,dsInsert])) ;
if TChancell.State=dsBrowse then FListe.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TChancell.BOF;
   DnEnable := Enabled and not TChancell.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
  (* if ((TUS_LIBELLE.AsString='US_LIBELLE') and (TUS_ABREGE.AsString='')) then
      TUS_ABEREGE.AsString:=Copy(Field.AsString,1,17) ;*)
   { Pas d'Abrégégé }
   END ;
end;

Procedure TFChancel.ChargeEnreg ;
Var Quotite : Integer ;
BEGIN
InitCaption(Self,CodeDevise.Value,CodeDevise.text) ;
Quotite:=QuotiteChange ;
TQuotite.Visible:=Quotite<>1 ; TValeurQ.Visible:=Quotite<>1 ;
if MonnaieIn then
  BEGIN
  TValable.Caption:=MsgBox.Mess[10]+'  '+DateToStr(V_PGI.DateDebutEuro-1)+'  '+MsgBox.Mess[12] ;
  END else
  BEGIN
  if SurEuro then TValable.Caption:=MsgBox.Mess[11]+'  '+DateToStr(V_PGI.DateDebutEuro)+'  '+MsgBox.Mess[13]
             else TValable.Caption:=MsgBox.Mess[10]+'  '+DateToStr(V_PGI.DateDebutEuro-1)+'  '+MsgBox.Mess[12] ;
  END ;
TValeurQ.Caption:=''+IntToStr(Quotite) ;
if Action=taConsult then BEGIN FicheReadOnly(Self) ; BPurge.Enabled:=False ; END ;
END ;

Function TFChancel.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve then Exit ;
   nbPost           : if Not EnregOK then exit ;
   nbDelete         : if MsgBox.execute(1,'','')<>mrYes then exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[7]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFChancel.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TChancell.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFChancel.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TChancell.state in [dsEdit,dsInsert]=False then Exit ;
if TChancell.state in [dsInsert] then
   BEGIN
   if TChancellH_DATECOURS.AsString='' then
      BEGIN
      MsgBox.Execute(5,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ;
      END ;
   if ChercheEnreg(TChancellH_DATECOURS.AsDateTime) then
      BEGIN
      MsgBox.Execute(4,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ;
      END ;
   if Not LimiteDateEuro(TChancellH_DATECOURS.AsDateTime) then
      BEGIN
      if MonnaieIn then MsgBox.Execute(14,'','') else
       if SurEuro then MsgBox.Execute(15,'','') else MsgBox.Execute(14,'','') ;
      FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ;
      END ;
   END ;
if TChancell.state in [dsEdit,dsInsert] then
   BEGIN
   if FListe.Columns[2].Visible then
      BEGIN
       if TChancellH_TAUXREEL.AsString='' then
          BEGIN
          MsgBox.Execute(6,'','') ; FListe.SelectedIndex:=2 ; FListe.SetFocus ; Exit ;
          END ;
       if TChancellH_TAUXREEL.AsFloat<=0 then
          BEGIN
          MsgBox.Execute(9,'','') ; FListe.SelectedIndex:=2 ; FListe.SetFocus ; Exit ;
          END ;
       END else
       BEGIN
       if TChancellH_COTATION.AsString='' then
          BEGIN
          MsgBox.Execute(6,'','') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; Exit ;
          END ;
       if TChancellH_COTATION.AsFloat<=0 then
          BEGIN
          MsgBox.Execute(9,'','') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; Exit ;
          END ;
       END ;
   END ;
Result:=TRUE  ;
END ;

Procedure TFChancel.NewEnreg ;
BEGIN
InitNew(TChancell) ;
TChancellH_DEVISE.AsString:=CodeDevise.Value ;
if(Double(FirstDate)>0)then TChancellH_DATECOURS.AsDateTime:=FirstDate
                       else TChancellH_DATECOURS.AsDateTime:=V_PGI.DateEntree ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.Setfocus ;
END ;

{Function TFChancel.DateRecente(DateOk : TDateTime) : Boolean ;
var Compris : TDateTime ;
    a, m, j    : word ;
BEGIN
// Verif de la Date
Result:=False ;
DecodeDate(Date,a,m,j) ;
a:=a-6 ; Compris:=EncodeDate(a,m,j) ;
if DateOk>Compris then
   BEGIN
   DecodeDate(Date,a,m,j) ;
   a:=a+104 ; Compris:=EncodeDate(a,m,j) ;
   if DateOk<Compris then Result:=True ;
   END ;
END ;}

Function TFChancel.ChercheEnreg(QuelDate : TDateTime) : boolean ;
BEGIN
Result:=ExisteSQL('SELECT H_DEVISE,H_DATECOURS FROM CHANCELL WHERE H_DEVISE="'+CodeDevise.Value+'" '+
                  'AND H_DATECOURS="'+UsDateTime(QuelDate)+'" ') ;
END ;

Function TFChancel.LimiteDateEuro(QuelDate : TDateTime) : boolean ;
BEGIN
Result:=False ;
// si monnaie in -> parité fixe
// sinon parité / Pivot avant DateDebutEuro
if MonnaieIn or (Not MonnaieIn and not SurEuro) then Result:=QuelDate<V_PGI.DateDebutEuro else
// ou parité / Euro après DateDebutEuro
  if (not MonnaieIn and SurEuro) then Result:=QuelDate>=V_PGI.DateDebutEuro ;
END ;

Function TFChancel.QuotiteChange : Integer ;
Var Q : TQuery ;
BEGIN
Result:=1 ;
Q:=OpenSQL('SELECT D_QUOTITE FROM DEVISE WHERE D_DEVISE="'+CodeDevise.Value+'"',TRUE,-1,'',true);
If Not Q.Eof Then Result:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
END ;

procedure TFChancel.TChancellH_TAUXREELChange(Sender: TField);
begin
if TChancellH_TAUXREEL.Value<0 then TChancellH_TAUXREEL.value:=0 ;
end;

procedure TFChancel.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFChancel.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFChancel.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFChancel.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFChancel.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFChancel.FormCreate(Sender: TObject);
begin PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; Lequel:='' ; end;

procedure TFChancel.BImprimerClick(Sender: TObject);
var MyBookmark: TBookmark;
begin
MyBookmark :=TChancell.GetBookmark ;
FListe.Columns.Items[1].Alignment:=taCenter ;
FListe.Columns.Items[2].Alignment:=taCenter ;
PrintDBGrid (FListe,PDevise,Caption,'') ;
FListe.Columns.Items[1].Alignment:=taRightJustify ;
FListe.Columns.Items[2].Alignment:=taRightJustify ;
TChancell.GotoBookmark(MyBookmark) ; TChancell.FreeBookmark(MyBookmark);
end;

procedure TFChancel.OkDateClick(Sender: TObject);
begin
H_DATEDU.Visible:=OkDate.Checked ;
H_DATEAU.Visible:=OkDate.Checked ;
TH_DATEDU.Visible:=OkDate.Checked ;
TH_DATEAU.Visible:=OkDate.Checked ;
ChangeIndex ;
end;

procedure TFChancel.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TChancell.Eof)then BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete); Key:=0 ; END ;
end;

procedure TFChancel.TChancellNewRecord(DataSet: TDataSet);
begin NewEnreg ; end;

procedure TFChancel.TChancellPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
begin
if TChancell.State=dsInsert then
   BEGIN MsgBox.Execute(8,'','') ; Action:=daAbort ; END ;
end;

procedure TFChancel.FListeRowEnter(Sender: TObject);
begin if TChancell.State=dsInsert then FListe.SelectedIndex:=0 ; end;

procedure TFChancel.SChancellUpdateData(Sender: TObject);
begin
if FListe.Columns[2].Visible then
   BEGIN
   if(TChancellH_TAUXREEL.AsFloat=0)And(TChancellH_TAUXCLOTURE.AsFloat=0)And
     (TChancellH_COMMENTAIRE.AsString='') then BEGIN TChancell.Cancel ; Exit ; END ;
   if TChancellH_TAUXREEL.AsFloat<=0 then
      BEGIN
      MsgBox.Execute(9,'','') ; FListe.SelectedIndex:=2 ; FListe.SetFocus ; SysUtils.Abort ;
      END ;
   END else
   BEGIN
   if(TChancellH_COTATION.AsFloat=0)And(TChancellH_TAUXCLOTURE.AsFloat=0)And
     (TChancellH_COMMENTAIRE.AsString='') then BEGIN TChancell.Cancel ; Exit ; END ;
   if TChancellH_COTATION.AsFloat<=0 then
      BEGIN
      MsgBox.Execute(9,'','') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; SysUtils.Abort ;
      END ;
   END ;
end;

procedure TFChancel.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFChancel.H_DATEDUExit(Sender: TObject);
begin
if ((TControl(Sender).Name='H_DATEDU') and (SurEuro)) then H_DATEDU.Text:=DateToStr(V_PGI.DateDebutEuro) else
 if ((TControl(Sender).Name='H_DATEAU') and (Not SurEuro)) then H_DATEAU.Text:=DateToStr(V_PGI.DateDebutEuro-1) ;
ChangeIndex ;
end;

procedure TFChancel.H_DATEAUExit(Sender: TObject);
begin
if ((Not SurEuro) and (StrToDate(H_DATEAU.Text)>V_PGI.DateDebutEuro)) then H_DATEAU.Text:=DateToStr(V_PGI.DateDebutEuro) ;
ChangeIndex ;
end;

procedure TFChancel.GereSens ( ForceCol : boolean ) ;
BEGIN
if TChancell.State in [dsEdit,dsInsert] then Bouge(nbPost) ;
if ForceCol then FListe.SelectedIndex:=3 ;
FListe.Columns[1].ReadOnly:=(RGSens.ItemIndex=1) ;
FListe.Columns[2].ReadOnly:=Not (FListe.Columns[1].ReadOnly) ;
FListe.Columns[1].Visible:=Not FListe.Columns[1].ReadOnly ;
FListe.Columns[2].Visible:=Not FListe.Columns[2].ReadOnly ;
END ;

procedure TFChancel.RGSensClick(Sender: TObject);
begin
GereSens(True) ;
end;

procedure TFChancel.TChancellBeforePost(DataSet: TDataSet);
begin
if Not SurEuro then Exit ;
if MonnaieIN then Exit ;
if FListe.Columns[1].Visible then
   BEGIN
   if Arrondi(TChancellH_Cotation.AsFloat,9)<>0 then TChancellH_TauxReel.AsFloat:=Arrondi(1/TChancellH_Cotation.AsFloat,9) else TChancellH_TauxReel.AsFloat:=0 ;
   END else
   BEGIN
   if Arrondi(TChancellH_TauxReel.AsFloat,9)<>0 then TChancellH_Cotation.AsFloat:=Arrondi(1/TChancellH_TauxReel.AsFloat,9) else TChancellH_Cotation.AsFloat:=0 ;
   END ;
end;

procedure TFChancel.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.
