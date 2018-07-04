unit Edtlegal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, Hctrls, ComCtrls, Spin,
  Mask, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, HEnt1, Ent1, ParamDBG, HDB, SaisUtil,
  Menus, HSysMenu, ParamDat, hmsgbox, HTB97
  ,HPanel, UIUtil // MODIF PACK AVANCE pour gestion mode inside
  , UObjFiltres, ADODB {JP 21/01/05 : FQ 15255}
  ;

Procedure VisuEditionLegale(Obligatoire : Boolean) ;
Procedure MajEdition(LeType,Exo,D1,D2,Dest : String ; Cd,Cc,Sd,Sc : Extended) ;
Procedure MajEditionLegal(LeType,Exo,D1,D2,Dest : String ; Cd,Cc,Sd,Sc : Extended) ;

type
  TFEdtLegal = class(TForm)
    Pages      : TPageControl;
    PCritere   : TTabSheet;
    SEdtl      : TDataSource;
    QEdtl      : THQuery;
    ECaption   : TEdit;
    TED_TYPEEDITION: THLabel;
    ED_TYPEEDITION: THValComboBox;
    TED_EXERCICE: THLabel;
    ED_EXERCICE: THValComboBox;
    Bevel1: TBevel;
    FindDialog: TFindDialog;
    HMTrad: THSystemMenu;
    XX_WHERE: TEdit;
    ECaption2: TEdit;
    TED_UTILISATEUR: THLabel;
    ED_UTILISATEUR: THValComboBox;
    TED_DATEEDITION: THLabel;
    ED_DATEEDITION: THCritMaskEdit;
    TTED_DATEEDITION: THLabel;
    ED_DATEEDITION_: THCritMaskEdit;
    ED_OBLIGATOIRE: TCheckBox;
    MsgBox: THMsgBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    Dock972: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    BCherche: TToolbarButton97;
    FFiltres: THValComboBox;
    PanelBouton: TToolWindow97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BParamListe: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BPurge: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BFermer: TToolbarButton97;
    BAide: TToolbarButton97;
    FLISTE: THDBGrid;
    procedure BFermerClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BPurgeClick(Sender: TObject);
    procedure SEdtlDataChange(Sender: TObject; Field: TField);
    procedure QEdtlAfterOpen(DataSet: TDataSet);
    procedure ED_DATEEDITIONKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    ObjFiltre : TObjFiltre; {JP 21/01/05 : Gestion des filtres FQ 15255}
    FILEDT    : String ;
    FindFirst : Boolean ;
    QR2ED_PERIODE : TDateTimeField ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
  public
    procedure ApresChargementFiltre;{JP 24/01/05 : FQ 15255}
  end;


implementation

{$R *.DFM}
uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  PrintDBG, Purge ;

procedure TFEdtLegal.BChercheClick(Sender: TObject);
begin QEdtL.UpdateCriteres ;  CentreDBGrid(FListe) ; end;

procedure TFEdtLegal.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFEdtLegal.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFEdtLegal.BParamListeClick(Sender: TObject);
begin ParamListe(QEdtl.Liste,FListe,QEdtl) ; end;

procedure TFEdtLegal.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
ED_DATEEDITION.text:=StDate1900 ; ED_DATEEDITION_.text:=StDate2099 ;
if ED_OBLIGATOIRE.State=cbChecked then
   BEGIN
   QEdtl.Liste:='EDTLEGALE' ; Caption:=Caption+' '+TraduireMemoire(ECaption.Text) ;
   FILEDT:='VISUEDTLEG' ; ED_TYPEEDITION.Datatype:='ttEditionLegale' ;
   HelpContext:=7709000 ;
   END else
   BEGIN
   QEdtl.Liste:='EDITION' ; Caption:=TraduireMemoire(ECaption2.Text) ;
   FILEDT:='VISUEDT' ; ED_TYPEEDITION.Datatype:='ttEdition' ;
   HelpContext:=7490000 ;
   END ;

  {JP 21/01/05 : Gestion des filtres FQ 15255}
  ObjFiltre.FFI_TABLE := FILEDT;
  ObjFiltre.Charger;

QEdtl.Manuel:=FALSE ; QEdtl.UpdateCriteres ;
BPurge.Visible:=(ED_OBLIGATOIRE.State<>cbChecked) ;
Pages.ActivePage:=PCritere ; CentreDBGrid(FListe) ;
UpdateCaption(Self) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion
Suite ........ : mode inside
Mots clefs ... :
*****************************************************************}
Procedure VisuEditionLegale(Obligatoire : Boolean) ;
var FEdtLegal:TFEdtLegal ;
    PP : THPanel ;
BEGIN
{$IFDEF CCS3}
  Exit ;
{$ENDIF}
  FEdtLegal:=TFEdtLegal.Create(Application) ;
  FEdtLegal.QEdtl.Manuel:=TRUE ;
  if Obligatoire
    then FEdtLegal.ED_OBLIGATOIRE.State:=cbChecked
    else FEdtLegal.ED_OBLIGATOIRE.State:=cbUnChecked ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FEdtLegal.ShowModal ;
      Finally
      FEdtLegal.Free ;
      End ;
    end
  else
    begin
    InitInside(FEdtLegal,PP) ;
    FEdtLegal.Show ;
    end ;

  Screen.Cursor:=crDefault ;
END ;

Procedure MajEdition(LeType,Exo,D1,D2,Dest : String ; Cd,Cc,Sd,Sc : Extended) ;
Var Q,Q1 : TQuery ;
    Legale : Boolean ;
    VirtExo : TExoDate ;
    D1Us,D2Us,DjUs,NomUser,St,St1 : String ;
    SCd,SCc,SSd,SSc : String ;
    M,A,NbrMois : Word ;
    DFMois,LeD1,LeD2 : TDateTime ;
    i,Num : Integer ;
BEGIN
{$IFDEF CCS3}
Exit ;
{$ENDIF}
D1Us:=UsDateTime(StrToDateTime(D1)) ; D2Us:=UsDateTime(StrToDateTime(D2)) ;
DjUs:=UsDateTime(Date) ; NomUser:=V_PGI.User ; Num:=1 ;
Q:=Tquery.Create(Application) ; Q.DataBaseName:='SOC' ; Q.Sql.Clear ;
Q1:=Tquery.Create(Application) ; Q1.DataBaseName:='SOC' ; Q1.Sql.Clear ;
St1:='INSERT INTO EDTLEGAL (ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_DATE1, '+
     'ED_DATE2, ED_DATEEDITION, ED_NUMEROEDITION, ED_CUMULDEBIT, ED_CUMULCREDIT, '+
     'ED_CUMULSOLDED, ED_CUMULSOLDEC, ED_UTILISATEUR, ED_DESTINATION) '+
     'VALUES ("-", "'+LeType+'", "'+Exo+'", "'+D1Us+'", "'+D2Us+'", "'+DjUs+'", '+
              IntToStr(Num)+', '+StrFPoint(Cd)+', '+StrFPoint(Cc)+', '+
              StrFPoint(Sd)+', '+StrFPoint(Sc)+', "'+NomUser+'", "'+Dest+'")' ;
Q1.Sql.Add(St1) ; ChangeSql(Q1) ; //Q1.Prepare ;
PrepareSQLODBC(Q1) ;
Q1.ExecSql ; Q1.Close ;

St:='Select CO_ABREGE From COMMUN Where CO_TYPE="EDL" And CO_CODE="'+LeType+'"' ;
Q.Sql.Add(St) ; ChangeSql(Q) ; Q.Open ;
Legale:=(Q.Fields[0].AsString='X') ; Q.Close ;
if Legale then
  BEGIN
  BeginTrans ;
  Cd:=0 ; Cc:=0 ; Sd:=0 ; Sc:=0 ;
  SCd:=StrfPoint(Cd) ; SCc:=StrfPoint(Cc) ; SSd:=StrfPoint(Sd) ; SSc:=StrfPoint(Sc) ;

  DFMois:=DebutDeMois(StrToDate(D1)) ;
  if(DFMois<StrTodate(D1)) then DFMois:=PlusMois(DFMois,1) ; LeD1:=DFMois ;

  DFMois:=FinDeMois(StrToDate(D2)) ;
  if(DFMois>StrTodate(D2)) then DFMois:=PlusMois(DFMois,-1) ; LeD2:=DFMois ;

  VirtExo.Deb:=LeD1 ; VirtExo.Fin:=LeD2 ; NombrePerExo(VirtExo,M,A,NbrMois) ;
  St:='Select ED_NUMEROEDITION From EDTLEGAL Where ED_TYPEEDITION="'+LeType+'" '+
      'AND ED_EXERCICE="'+Exo+'" AND ED_OBLIGATOIRE="X" AND ED_PERIODE=:Per' ;
  Q.Close ; Q.Sql.Clear ; Q.Sql.Add(St) ; ChangeSql(Q) ; //Q.Prepare ;
  PrepareSQLODBC(Q) ;
  for i:=1 to NbrMois do
    BEGIN
     D1Us:=UsDateTime(LeD1) ; D2Us:=UsDateTime(FinDeMois(LeD1)) ;

     Q.Close ; Q.Params[0].AsDateTime:=LeD1 ; Q.Open ;
     FetchSqlODBC(Q) ;
     if Q.Eof then Num:=1 else Num:=Q.Fields[0].AsInteger+1 ;

     if Num=1 then
        St1:='INSERT INTO EDTLEGAL (ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_DATE1,'+
             'ED_DATE2, ED_DATEEDITION, ED_NUMEROEDITION, ED_PERIODE, ED_CUMULDEBIT, '+
             'ED_CUMULCREDIT, ED_CUMULSOLDED, ED_CUMULSOLDEC, '+
             'ED_UTILISATEUR, ED_DESTINATION)'+
             ' VALUES ("X", "'+LeType+'", "'+Exo+'", "'+D1Us+'", "'+D2Us+'", "'+DjUs+'", '+
                       IntToStr(Num)+', "'+D1Us+'", '+SCd+', '+SCc+', '+
                       SSd+', '+SSc+', "'+NomUser+'", "'+Dest+'")'
     else
        St1:='UPDATE EDTLEGAL SET ED_OBLIGATOIRE="X", ED_TYPEEDITION="'+LeType+'", '+
             'ED_EXERCICE="'+Exo+'", ED_DATE1="'+D1Us+'", ED_DATE2="'+D2Us+'", ED_DATEEDITION="'+DjUs+'", '+
             'ED_NUMEROEDITION='+IntToStr(Num)+', ED_PERIODE="'+D1Us+'", ED_CUMULDEBIT='+SCd+', '+
             'ED_CUMULCREDIT='+SCc+', ED_CUMULSOLDED='+SSd+', '+
             'ED_CUMULSOLDEC='+SSc+', ED_UTILISATEUR="'+NomUser+'", ED_DESTINATION="'+Dest+'" '+
             'Where ED_TYPEEDITION="'+LeType+'" AND ED_EXERCICE="'+Exo+'" '+
             'AND ED_OBLIGATOIRE="X" AND ED_PERIODE="'+D1Us+'"' ;
     Q.Close ; ExecuteSQL(St1) ;
//     Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(St1) ; ChangeSql(Q1) ; Q1.Prepare ; Q1.ExecSql ;
     LeD1:=PlusMois(LeD1,1) ;
    END ;
  CommitTrans ;
  END ;
Ferme(Q) ; Ferme(Q1) ;
END ;

Procedure MajEditionLegal(LeType,Exo,D1,D2,Dest : String ; Cd,Cc,Sd,Sc : Extended) ;
Var Q,Q1 : TQuery ;
    Legale : Boolean ;
    VirtExo : TExoDate ;
    D1Us,D2Us,DjUs,NomUser,St,St1 : String ;
    SCd,SCc,SSd,SSc : String ;
    M,A,NbrMois : Word ;
    DFMois,LeD1,LeD2 : TDateTime ;
    i,Num : Integer ;
BEGIN
{$IFDEF CCS3}
Exit ;
{$ENDIF}
D1Us:=UsDateTime(StrToDateTime(D1)) ; D2Us:=UsDateTime(StrToDateTime(D2)) ;
DjUs:=UsDateTime(Date) ; NomUser:=V_PGI.User ;
Q:=Tquery.Create(Application) ; Q.DataBaseName:='SOC' ; Q.Sql.Clear ;
St:='Select CO_ABREGE From COMMUN Where CO_TYPE="EDL" And CO_CODE="'+LeType+'"' ;
Q.Sql.Add(St) ; ChangeSql(Q) ; Q.Open ;
Q1:=Tquery.Create(Application) ; Q1.DataBaseName:='SOC' ;
Legale:=(Q.Fields[0].AsString='X') ; Q.Close ;
if Legale then
  BEGIN
  BeginTrans ;
  Cd:=0 ; Cc:=0 ; Sd:=0 ; Sc:=0 ;
  SCd:=StrfPoint(Cd) ; SCc:=StrfPoint(Cc) ; SSd:=StrfPoint(Sd) ; SSc:=StrfPoint(Sc) ;

  DFMois:=DebutDeMois(StrToDate(D1)) ;
  if(DFMois<StrTodate(D1)) then DFMois:=PlusMois(DFMois,1) ; LeD1:=DFMois ;

  DFMois:=FinDeMois(StrToDate(D2)) ;
  if(DFMois>StrTodate(D2)) then DFMois:=PlusMois(DFMois,-1) ; LeD2:=DFMois ;

  VirtExo.Deb:=LeD1 ; VirtExo.Fin:=LeD2 ; NombrePerExo(VirtExo,M,A,NbrMois) ;
  St:='Select ED_NUMEROEDITION From EDTLEGAL Where ED_TYPEEDITION="'+LeType+'" '+
      'AND ED_EXERCICE="'+Exo+'" AND ED_OBLIGATOIRE="X" AND ED_PERIODE=:Per' ;
  Q.Close ; Q.Sql.Clear ; Q.Sql.Add(St) ; ChangeSql(Q) ; //Q.Prepare ;
  PrepareSQLODBC(Q) ;
  for i:=1 to NbrMois do
    BEGIN
     D1Us:=UsDateTime(LeD1) ; D2Us:=UsDateTime(FinDeMois(LeD1)) ;

     Q.Close ; Q.Params[0].AsDateTime:=LeD1 ; Q.Open ;
     FetchSqlODBC(Q) ;
     if Q.Eof then Num:=1 else Num:=Q.Fields[0].AsInteger+1 ;

     if Num=1 then
        St1:='INSERT INTO EDTLEGAL (ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_DATE1,'+
             'ED_DATE2, ED_DATEEDITION, ED_NUMEROEDITION, ED_PERIODE, ED_CUMULDEBIT, '+
             'ED_CUMULCREDIT, ED_CUMULSOLDED, ED_CUMULSOLDEC, '+
             'ED_UTILISATEUR, ED_DESTINATION)'+
             ' VALUES ("X", "'+LeType+'", "'+Exo+'", "'+D1Us+'", "'+D2Us+'", "'+DjUs+'", '+
                       IntToStr(Num)+', "'+D1Us+'", '+SCd+', '+SCc+', '+
                       SSd+', '+SSc+', "'+NomUser+'", "'+Dest+'")'
     else
        St1:='UPDATE EDTLEGAL SET ED_OBLIGATOIRE="X", ED_TYPEEDITION="'+LeType+'", '+
             'ED_EXERCICE="'+Exo+'", ED_DATE1="'+D1Us+'", ED_DATE2="'+D2Us+'", ED_DATEEDITION="'+DjUs+'", '+
             'ED_NUMEROEDITION='+IntToStr(Num)+', ED_PERIODE="'+D1Us+'", ED_CUMULDEBIT='+SCd+', '+
             'ED_CUMULCREDIT='+SCc+', ED_CUMULSOLDED='+SSd+', '+
             'ED_CUMULSOLDEC='+SSc+', ED_UTILISATEUR="'+NomUser+'", ED_DESTINATION="'+Dest+'" '+
             'Where ED_TYPEEDITION="'+LeType+'" AND ED_EXERCICE="'+Exo+'" '+
             'AND ED_OBLIGATOIRE="X" AND ED_PERIODE="'+D1Us+'"' ;
     Q.Close ; ExecuteSQL(St1) ;
//     Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(St1) ; ChangeSql(Q1) ; Q1.Prepare ; Q1.ExecSql ;
     LeD1:=PlusMois(LeD1,1) ;
    END ;
  CommitTrans ;
  END ;
Ferme(Q) ; Ferme(Q1) ;
END ;

procedure TFEdtLegal.BRechercherClick(Sender: TObject);
begin FindFirst:=True; FindDialog.Execute ; end;

procedure TFEdtLegal.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FindFirst); end;

procedure TFEdtLegal.BImprimerClick(Sender: TObject);
begin PrintDBGrid(FListe,Pages,Caption,'') ; end;

procedure TFEdtLegal.BFermerClick(Sender: TObject);
begin Close ; end;

procedure TFEdtLegal.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFEdtLegal.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, '');
  ObjFiltre.ApresChangementFiltre := ApresChargementFiltre;

  WMinX:=Width ;
  WMinY:=320 ;
end;

procedure TFEdtLegal.BPurgeClick(Sender: TObject);
Var LesDatesASup : TDateTime ;
    StPlus : String ;
begin
// ED_OBLIGATOIRE, ED_TYPEEDITION, ED_EXERCICE, ED_PERIODE
if PurgeOui(StDate1900,MsgBox.Mess[0],LesDatesASup) then
   BEGIN
   StPlus:='' ;
   If ED_TYPEEDITION.Value<>'' then StPlus:=' and ED_TYPEEDITION="'+ED_TYPEEDITION.Value+'" ' ;
   If ED_EXERCICE.Value<>'' then StPlus:=StPlus+' and ED_EXERCICE="'+ED_EXERCICE.Value+'" ' ;
   ExecuteSQl('DELETE FROM EDTLEGAL WHERE ED_OBLIGATOIRE<>"X" and ED_TYPEEDITION<>"TIF" '+
              ' AND ED_DATEEDITION<="'+USDatetime(LesDatesASup)+'" '+StPlus+' ') ;
   QEdtl.Close ; QEdtl.Open ;
   END ;
end;

procedure TFEdtLegal.SEdtlDataChange(Sender: TObject; Field: TField);
begin
BPurge.Enabled:=Not (QEdtl.Eof and QEdtl.Bof) ;
end;

procedure TFEdtLegal.QEdtlAfterOpen(DataSet: TDataSet);
begin
QR2ED_PERIODE :=TDateTimeField(QEdtl.FindField('ED_PERIODE')) ;
If QR2ED_PERIODE<>Nil then QR2ED_PERIODE.Displayformat:='mmmm-yyyy' ;
end;

procedure TFEdtLegal.ED_DATEEDITIONKeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self, Sender, Key) ;
end;

procedure TFEdtLegal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{---------------------------------------------------------------------------------------}
procedure TFEdtLegal.FormClose(Sender: TObject; var Action: TCloseAction);
{---------------------------------------------------------------------------------------}
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  FreeAndNil(ObjFiltre);
end;

{---------------------------------------------------------------------------------------}
procedure TFEdtLegal.ApresChargementFiltre;
{---------------------------------------------------------------------------------------}
begin
  {JP 24/01/05 : On lance la recherche après chargement du filtre}
  if FFiltres.Text <> '' then BChercheClick(BCherche);
end;

end.
