{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 26/04/2002
Modifié le ... : 26/04/2002
Description .. : Depuis 04/2002, passage en E-AGL :
Suite ........ :  - Passage de la fiche dans Décla "CPCHPLIBRELIB",
Suite ........ :  - Création d'une unité TOF : CPCHPLIBRELIB_TOF,
Mots clefs ... : CHAMPS LIBRES
*****************************************************************}
unit LibChpLi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics,Buttons,HSysMenu, Grids,
  Controls, Forms, Dialogs, StdCtrls, ComCtrls,ExtCtrls, hmsgbox,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DBGrids, HDB, PrintDBG, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBCtrls,
{$ENDIF}
  UtilPGI, Ent1, Hctrls,HEnt1,
	Utob, CPCHPLIBRELIB_TOF, utobdebug;

procedure AGLModifLibelleChampLibre(sTable,sLibTable : string) ;
procedure PersoChampsLibres(sTable : string ; t : TWinControl ; vBoAvecTOB : boolean = false ) ;
function  PersoChamp(sTable : string ; var s : String ; var vv : boolean) : boolean ;
function PersoChampAvecTOB(ParamLibTOB : TOB; sTable : string ; var s : String ; var vv : boolean) : boolean ;

{$IFNDEF EAGLCLIENT}
procedure ModifLibelleChampLibre(sTable,sLibTable : string) ;
type
  TFLibChpLi = class(TForm)
    HPB: TPanel;
    BAnnuler: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    Q: TQuery;
    DS: TDataSource;
    PTop: TPanel;
    FListe: THDBGrid;
    FAutoSave: TCheckBox;
    MsgBox: THMsgBox;
    DBNav: TDBNavigator;
    rTMvt: TRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFermeClick(Sender: TObject);
    procedure DSDataChange(Sender: TObject; Field: TField);
    procedure QNewRecord(DataSet: TDataSet);
    procedure BImprimerClick(Sender: TObject);
    procedure rTMvtClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TypeTable : String ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    procedure InverseSelection ;
    function  FiltreTable : string ;
  public
    { Déclarations publiques }
  end;
{$ENDIF}

implementation

{$R *.DFM}

//==============================================================================
procedure AGLModifLibelleChampLibre(sTable,sLibTable : string) ;
var
	Args : String ;
BEGIN
  if sTable <> ''
  	then Args := sTable +';' + sLibTable
		else Args := '';
	CPLanceFiche_ParamChampsLibres(Args);
END ;
//==============================================================================

function PersoChamp(sTable : string ; var s : String ; var vv : boolean) : boolean ;
var Q : TQuery ;
    SQL : string ;
begin
  result := False ;
  SQL := 'SELECT PL_LIBELLE, PL_VISIBLE FROM PARAMLIB ' +
  			 'WHERE PL_TABLE="' + sTable + '" AND PL_CHAMP="' + s + '"' ;
  Q := OpenSQL(SQL,True) ;
  if not Q.EOF then
     begin
     s := Q.FindField('PL_LIBELLE').AsString ;
     if s='' then s := Q.FindField('PL_LIBELLE').AsString ;
     vv := (Q.FindField('PL_VISIBLE').AsString='X') ;
     result := True ;
     end ;
  Ferme(Q) ;
end ;

procedure PersoChampsLibres(sTable : string ; t : TWinControl ; vBoAvecTOB : boolean = false ) ;
var i,n  : integer ;
    C,CC : TControl ;
    s    : string ;
    v    : boolean ;
    lTOB : TOB ;
    lQ   : TQuery ;
    lBoR : boolean ;
begin
n := 0 ; lTOB := nil ;
if vBoAvecTOB  then
 begin

  lQ := OpenSQL('SELECT PL_CHAMP, PL_LIBELLE, PL_VISIBLE FROM PARAMLIB ' +
                'WHERE PL_TABLE="' + sTable + '" ' , true ) ;
  lTOB := TOB.Create('',nil,-1) ;
  // YMO 06/06/06 FQ17643 utilisation de 'LoadDetailDB' et non 'SelectDb' pour ramener ts les enregs
  lTOB.LoadDetailDB('', '', '',lQ, False ) ;
  Ferme(lQ) ;
  
 end ; // if
for i:=0 to t.ControlCount-1 do
   begin 
   C := t.Controls[i] ;
   if C is THLabel then
      begin
      CC := TControl(THLabel(C).FocusControl) ;
      if CC=nil then continue ;
      s := CC.Name ;
      if s[1]='_' then s:=sTable+s ;
      if vBOAvecTOB then lBoR := PersoChampAvecTOB(lTOB,STable,s,v) else lBoR := PersoChamp(sTable,s,v) ;
      if lBoR then
         begin
         C.Enabled := v ;
         CC.Enabled := v ;
         THLabel(C).Caption := s ;
         if v then Inc(n) ;
         end ;
      end else
   if C is TCheckBox then
      begin
      s := TCheckBox(C).Name ;
      if s[1]='_' then s:=sTable+s ;
      if vBOAvecTOB then lBoR := PersoChampAvecTOB(lTOB,STable,s,v) else lBoR := PersoChamp(sTable,s,v) ;
      if lBoR then
         begin
         C.Enabled := v ;
         TCheckBox(C).Caption := s ;
         if v then Inc(n) ;
         end ;
      end else
{$IFNDEF EAGLCLIENT}
   if C is TDBCheckBox then
      begin
      s := TDBCheckBox(C).DataField ;
      if vBOAvecTOB then lBoR := PersoChampAvecTOB(lTOB,STable,s,v) else lBoR := PersoChamp(sTable,s,v) ;
      if lBoR then
         begin
         C.Enabled := v ;
         TDBCheckBox(C).Caption := s ;
         if v then Inc(n) ;
         end ;
      end ;
{$ENDIF}
   end ;
if n=0 then
   begin
   if t is TTabSheet then TTabSheet(t).TabVisible:=False
                     else t.Visible := False ;
   end ;

 lTOB.Free ;

end ;



{$IFNDEF EAGLCLIENT} //=========================================> FICHE NON EAGL

procedure ModifLibelleChampLibre(sTable,sLibTable : string) ;
var FLibChpLi : TFLibChpLi ;
BEGIN
FLibChpLi:=TFLibChpLi.Create(Application) ;
 Try
  if sTable<>'' then
     begin
     FLibChpLi.Caption := sLibTable ;
     FLibChpLi.TypeTable := sTable ;
     FLibChpLi.PTop.Visible := False ;
     end else
     begin
     FLibChpLi.TypeTable := '' ;
     end ;
  FLibChpLi.ShowModal ;
 Finally
  FLibChpLi.Free ;
 End ;
SourisNormale ;
END ;

Function TFLibChpLi.EnregOK : boolean ;
var slib : string ;
BEGIN
Result:=False ;
//OldTrad if (V_PGI.Langue='FRA')  or (V_PGI.TradMemoire) then slib:='PL_LIBELLE' else slib:='PL__'+V_PGI.Langue ;
slib:='PL_LIBELLE' ;
if Q.state in [dsEdit,dsInsert]=False then Exit ;
if Q.state in [dsEdit] then
   begin
   if Q.FindField(slib).AsString='' then Q.FindField(slib).AsString:=Q.FindField('PL_LIBELLE').AsString ;
   if Q.FindField('PL_VISIBLE').AsString='' then Q.FindField('PL_VISIBLE').AsString:='-' ;
   end ;
Result:=True ;
END ;

Function TFLibChpLi.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if Q.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,caption,'') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFLibChpLi.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[2]) ;
Result:=TRUE ;
END ;

procedure TFLibChpLi.InverseSelection ;
begin
if Q.State<>dsEdit then Q.Edit ;
if Q.FindField('PL_VISIBLE').AsString='-' then Q.FindField('PL_VISIBLE').AsString:='X'
                                          else Q.FindField('PL_VISIBLE').AsString:='-' ;
end ;

function TFLibChpLi.FiltreTable : string ;
var Q : TQuery ;
    n : integer ;
    sTable,sWhere,sChamp : string ;
begin
sWhere := '' ;
case rTMvt.ItemIndex of
   0 : sTable := 'E' ;
   1 : sTable := 'A' ;
   2 : sTable := 'U' ;
   end ;
Q := OpenSQL('SELECT * FROM COMMUN WHERE CC_TYPE="NAT" AND CC_CODE LIKE "'+ sTable + '%" ORDER BY CC_CODE',True) ;
n := 0 ;
while not Q.EOF do
   begin
   if Q.FindField('CC_ABREGE').AsString <> 'X' then
      begin
      sChamp := TypeTable + '_TABLE' + IntToStr(n) ;
      sWhere := sWhere + ' AND PL_CHAMP<>"' + sChamp + '"' ;
      end ;
   Q.Next ;
   Inc(n) ;
   end ;
Ferme(Q) ;
result := sWhere ;
end ;

{--- Evenèments de la form ---}

procedure TFLibChpLi.FormShow(Sender: TObject);
begin
rTMvt.ItemIndex := 0 ;
rTMvtClick(nil) ;
end;

procedure TFLibChpLi.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if FListe.SelectedField<>Q.FindField('PL_VISIBLE') then exit ;
if Key=' ' then InverseSelection ;
if (Key<>'-') and (Key<>'X') and (Key<>'x') then Key:=#0 ;
if Key='x' then Key:='X' ;
end;

procedure TFLibChpLi.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin if((ssCtrl in Shift) and (Key=VK_DELETE)) then Key:=0 ; end;

procedure TFLibChpLi.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) end;

procedure TFLibChpLi.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFLibChpLi.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFLibChpLi.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFLibChpLi.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFLibChpLi.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFLibChpLi.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFLibChpLi.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFLibChpLi.DSDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
BEGIN
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not Q.BOF;
   DnEnable := Enabled and not Q.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   END ;
END ;

procedure TFLibChpLi.QNewRecord(DataSet: TDataSet);
begin Bouge(nbCancel) ; end;

procedure TFLibChpLi.BImprimerClick(Sender: TObject);
begin PrintDBGrid(FListe,Nil,VireSouligne(Msg.Mess[5]+' '+rTmvt.Items[rTmvt.ItemIndex]),''); end;

procedure TFLibChpLi.rTMvtClick(Sender: TObject);
begin
if PTop.Visible then
   begin
   case rTMvt.ItemIndex of
      0 : TypeTable := 'E' ;
      1 : TypeTable := 'Y' ;
      2 : TypeTable := 'BE' ;
      end ;
   end ;
if Q.Active then Q.Active:=False ;
Q.SQL.Clear ;
Q.SQL.Add('SELECT * FROM PARAMLIB WHERE PL_TABLE="' + TypeTable + '" ' + FiltreTable) ;
ChangeSQL(Q) ;
(*OldTrad
if (V_PGI.Langue='FRA')  or (V_PGI.TradMemoire) then FListe.Columns.Items[0].FieldName:='PL_LIBELLE'
                                          else FListe.Columns.Items[0].FieldName:='PL__'+V_PGI.Langue ;
*)
Q.Open ;
end;

procedure TFLibChpLi.FListeDblClick(Sender: TObject);
begin
InverseSelection ;
end;

procedure TFLibChpLi.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{$ENDIF} //==========================================================> FIN FICHE

//==============================================================================

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 02/04/2002
Modifié le ... : 26/04/2002
Description .. : Meme usage que PersoChamp.
Suite ........ : Le paramètre ParamLibTOB doit être une TOB virtuelle
Suite ........ : contenant un ensemble des enregistrements de ParamLib
Suite ........ : Ex :
Suite ........ :    ParamTOB.Create('Paramètres Tables Libres',nil,-1);
Suite ........ :    Q := OpenSQL('select * from paramlib where Pl_table
Suite ........ : ="E"');
Suite ........ :    ParamTOB.loadDetailDB('PARAMLIB','','',Q,False);
Suite ........ :    Ferme(Q);
Suite ........ :
Suite ........ : Evite ainsi l'exécution de X requêtes pour un appel dans
Suite ........ : une boucle
Mots clefs ... : PERSOCHAMP PERSOCHAMPAVECTOB
*****************************************************************}
function PersoChampAvecTOB(ParamLibTOB : TOB; sTable : string ; var s : String ; var vv : boolean) : boolean ;
var
	TOBFille : TOB;
begin
  result := False ;
  // YMO 06/06/06 FQ17643 Rajout PL_CHAMP, utilisé pour test dans 'PersoChampAvecTOB'
  TOBFille := ParamLibTOB.FindFirst(['PL_CHAMP'],[s],False);

  if TOBFille <> nil then
   	Begin
   	s := TOBFille.GetValue('PL_LIBELLE');
   	vv := TOBFille.GetValue('PL_VISIBLE')='X' ;
   	result := True ;
   	End ;
end ;
//==============================================================================
end.
