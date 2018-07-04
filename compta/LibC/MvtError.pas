{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 02/03/2004
Description .. :
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... :
*****************************************************************}
unit MvtError;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB, DBCtrls, hmsgbox, Grids, DBGrids, HDB, StdCtrls, Buttons,
  ExtCtrls, Ent1, HEnt1, Hctrls , Hcompte, UTILEDT, RappType, SAISUTIL,
  HSysMenu;

Procedure ChercheError(Mvt : TInfoMvt) ;

type
  TFMvtError = class(TForm)
    PBouton: TPanel;
    FAutoSave: TCheckBox;
    FListe: THDBGrid;
    MsgBox: THMsgBox;
    DBNav: TDBNavigator;
    SEcr: TDataSource;
    QEcr: TQuery;
    TModif: TLabel;
    MsgLibel: THMsgBox;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    Panel2: TPanel;
    BZOOM: THBitBtn;
    HMTrad: THSystemMenu;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    BAnnuler: THBitBtn;
    BParamListe: THBitBtn;
    BRechercher: THBitBtn;
    FindDialog: TFindDialog;
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure SEcrStateChange(Sender: TObject);
    procedure SEcrUpdateData(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BZOOMClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure FListeColEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TrouveMvt  : TInfoMvt ;
    Modifier, FindFirst : boolean ;
    NomCompte  : String ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Function  chercheInfo : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Function  SiMvt : Boolean ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses
{$IFNDEF PGIIMMO}
Lettrage,
{$ENDIF}
PrintDBG,
CPGeneraux_TOM,
CPTiers_TOM,
CPSection_TOM,
CPJournal_TOM ;

Procedure ChercheError(Mvt : TInfoMvt) ;
var FMvtError: TFMvtError ;
BEGIN
FMvtError:=TFMvtError.Create(Application) ;
 Try
  FMvtError.TrouveMvt:=Mvt ;
  FMvtError.ShowModal ;
 Finally
  FMvtError.Free ;
 End ;
Screen.Cursor:=SynCrDefault ;
END ;

procedure TFMvtError.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Case TrouveMvt.indicetable of
   3 : BZOOM.hint:=MsgLibel.Mess[6] ;
   4 : BZOOM.hint:=MsgLibel.Mess[5] ;
  Else BZOOM.hint:='' ;
  end ;
BZOOM.Enabled:=(TrouveMvt.indicetable=3)or(TrouveMvt.indicetable=4) ;
TModif.Caption:='' ; Modifier:=False ;
if not chercheInfo then BEGIN MessageAlerte(MsgBox.Mess[5]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
end;

procedure TFMvtError.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

Function TFMvtError.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,
     nbnext,nbfirst : if Not OnSauve  then Exit ;
     nbPost         : if Not EnregOK  then Exit ;
    end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.mess[5]) ;
result:=TRUE ;
END ;

Function TFMvtError.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if QEcr.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.Execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
 mrYes : if not Bouge(nbPost) then exit ;
 mrNo  : if not Bouge(nbCancel) then exit ;
 mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFMvtError.EnregOK : boolean ;
BEGIN
result:=FALSE  ; Modifier:=True ;
TModif.Caption:=MsgLibel.Mess[0] ;
if QEcr.state in [dsEdit,dsInsert]=False then Exit ;
if QEcr.state in [dsEdit,dsInsert] then
   BEGIN
   END ;
TModif.Caption:=MsgLibel.Mess[1] ;
Result:=TRUE  ; Modifier:=False ;
END ;

procedure TFMvtError.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFMvtError.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFMvtError.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFMvtError.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFMvtError.BValiderClick(Sender: TObject);
begin Modifier:=False ;Bouge(nbPost) ; end;

procedure TFMvtError.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; FListe.ReadOnly:=True ; end;

procedure TFMvtError.BFermeClick(Sender: TObject);
begin Close ; end;



Function  TFMvtError.chercheInfo : boolean ;
Var Ok : Boolean ;
BEGIN
Ok:=False ;
Case TrouveMvt.indicetable of
   1 : BEGIN  { Cherche Piece Compta Ecriture}
       QEcr.Close ;
       QEcr.SQL.Clear ;
       QEcr.SQL.Add(' Select * ') ;
       QEcr.SQL.Add(' From ECRITURE ') ;
       QEcr.SQL.Add(' Where ') ;
       QEcr.SQL.Add(' E_EXERCICE="'+TrouveMvt.exercice+'" ') ;
       QEcr.SQL.Add(' And E_JOURNAL="'+TrouveMvt.journal+'"') ;
       QEcr.SQL.Add(' And E_DATECOMPTABLE="'+USDATETIME(TrouveMvt.datecomptable)+'"') ;
       QEcr.SQL.Add(' And E_NUMEROPIECE='+IntToStr(TrouveMvt.numeropiece)+'') ;
       QEcr.SQL.Add(' order by E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE,E_NUMLIGNE,E_DEVISE ') ;
       ChangeSql(QEcr) ; //QEcr.Prepare ;
       PrepareSQLODBC(QEcr) ;
       QEcr.Open ;
       Ok:=not QEcr.Eof ;
       END ;
   2 : BEGIN  { Cherche Piece Compta Analytique}
       QEcr.Close ;
       QEcr.SQL.Clear ;
       QEcr.SQL.Add(' Select * ') ;
       QEcr.SQL.Add(' From ANALYTIQ ') ;
       QEcr.SQL.Add(' Where ') ;
       QEcr.SQL.Add(' Y_EXERCICE="'+TrouveMvt.exercice+'" ') ;
       if TrouveMvt.AXEPLUS[1]<>'' then
          BEGIN
          QEcr.SQL.Add(' And (Y_AXE="'+TrouveMvt.AXEPLUS[1]+'" or Y_AXE="'+TrouveMvt.AXEPLUS[2]+'" or ') ;
          QEcr.SQL.Add(' Y_AXE="'+TrouveMvt.AXEPLUS[3]+'" or Y_AXE="'+TrouveMvt.AXEPLUS[4]+'" or Y_AXE="'+TrouveMvt.AXEPLUS[5]+'" ) ') ;
          END Else QEcr.SQL.Add(' And Y_AXE="'+TrouveMvt.axe+'" ') ;
       QEcr.SQL.Add(' And Y_JOURNAL="'+TrouveMvt.journal+'"') ;
       QEcr.SQL.Add(' And Y_DATECOMPTABLE="'+USDATETIME(TrouveMvt.datecomptable)+'"') ;
       QEcr.SQL.Add(' And Y_NUMEROPIECE='+IntToStr(TrouveMvt.numeropiece)+'') ;
       //QEcr.SQL.Add(' And E_NUMLIGNE='+IntToStr(TrouveMvt.numligne)+'') ;
       QEcr.SQL.Add(' order by Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE,Y_NUMLIGNE,Y_DEVISE ') ;
       ChangeSql(QEcr) ; //QEcr.Prepare ;
       PrepareSQLODBC(QEcr) ;
       QEcr.Open ;
       Ok:=not QEcr.Eof ;
       END ;
   3 : BEGIN  { Cherche Mvt Lettre ou non-Lettre}
       QEcr.Close ;
       QEcr.SQL.Clear ;
       QEcr.SQL.Add(' Select * ') ;
       QEcr.SQL.Add(' From ECRITURE ') ;
       QEcr.SQL.Add(' Where ') ;
       QEcr.SQL.Add(' E_LETTRAGE="'+TrouveMvt.lettrage+'" ') ;
       QEcr.SQL.Add(' And E_AUXILIAIRE="'+TrouveMvt.auxiliaire+'" ') ;
       QEcr.SQL.Add(' And E_GENERAL="'+TrouveMvt.general+'" ') ;
       QEcr.SQL.Add(' And E_ETATLETTRAGE="'+TrouveMvt.etatlettrage+'" ') ;
       QEcr.SQL.Add(' order by E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE,E_LETTRAGEDEV, E_DEVISE ') ;
       ChangeSql(QEcr) ; //QEcr.Prepare ;
       PrepareSQLODBC(QEcr) ;
       QEcr.Open ;
       Ok:=not QEcr.Eof ;
       END ;
   4 : BEGIN  { Cherche ligne compte général }
       QEcr.Close ;
       QEcr.SQL.Clear ;
       QEcr.SQL.Add('Select *') ;
       if TrouveMvt.NATUREPIECE='G' then QEcr.SQL.Add(' From GENERAUX where G_GENERAL="'+TrouveMvt.general+'" ') else
       if TrouveMvt.NATUREPIECE='T' then QEcr.SQL.Add(' From TIERS where T_AUXILIAIRE="'+TrouveMvt.auxiliaire+'" ') else
       if TrouveMvt.NATUREPIECE='S' then QEcr.SQL.Add(' From SECTION where S_AXE="'+TrouveMvt.axe+'" and S_SECTION="'+TrouveMvt.section+'" ') else
       if TrouveMvt.NATUREPIECE='J' then QEcr.SQL.Add(' From JOURNAL where J_JOURNAL="'+TrouveMvt.journal+'" ') ;
       ChangeSql(QEcr) ; //QEcr.Prepare ;
       PrepareSQLODBC(QEcr) ;
       QEcr.Open ;
       Ok:=not QEcr.Eof ;
       If TrouveMvt.NATUREPIECE='G' then
          BEGIN
          NomCompte:=QEcr.FindField('G_GENERAL').AsString+' '+QEcr.FindField('G_ABREGE').AsString ;
          TrouveMvt.Modif:=Not SiMvt ;
          Fliste.Columns[0].ReadOnly:=Not TrouveMvt.Modif ;   // Code
          Fliste.Columns[3].ReadOnly:=Not TrouveMvt.Modif ;   // Nature
          Fliste.Columns[9].ReadOnly:=Not TrouveMvt.Modif ;   // Collectif
          Fliste.Columns[25].ReadOnly:=Not TrouveMvt.Modif ;  // Lettrable
          Fliste.Columns[26].ReadOnly:=Not TrouveMvt.Modif ;  // Pointable
          Fliste.Columns[27].ReadOnly:=Not TrouveMvt.Modif ;  // Ventilable :
          Fliste.Columns[41].ReadOnly:=Not TrouveMvt.Modif ;  // V 1
          Fliste.Columns[42].ReadOnly:=Not TrouveMvt.Modif ;  // V 2
          Fliste.Columns[43].ReadOnly:=Not TrouveMvt.Modif ;  // V 3
          Fliste.Columns[44].ReadOnly:=Not TrouveMvt.Modif ;  // V 4
          Fliste.Columns[45].ReadOnly:=Not TrouveMvt.Modif ;  // V 5
          END Else
       If TrouveMvt.NATUREPIECE='T' then
          BEGIN
          NomCompte:=QEcr.FindField('T_AUXILIAIRE').AsString+' '+QEcr.FindField('T_ABREGE').AsString ;
          TrouveMvt.Modif:=Not SiMvt ;
          Fliste.Columns.Items[0].ReadOnly:=Not TrouveMvt.Modif ;   // Code Cpta
          Fliste.Columns.Items[1].ReadOnly:=Not TrouveMvt.Modif ;   // Nature
          Fliste.Columns.Items[4].ReadOnly:=Not TrouveMvt.Modif ;   // Collectif
          END Else
       If TrouveMvt.NATUREPIECE='J' then
          BEGIN
          NomCompte:=QEcr.FindField('J_JOURNAL').AsString+' '+QEcr.FindField('J_ABREGE').AsString ;
          TrouveMvt.Modif:=Not SiMvt ;
          Fliste.Columns.Items[0].ReadOnly:=Not TrouveMvt.Modif ;   // Code
          Fliste.Columns.Items[3].ReadOnly:=not TrouveMvt.Modif ;   // Nature
          Fliste.Columns.Items[45].ReadOnly:=Not TrouveMvt.Modif ;  // Axe
          Fliste.Columns.Items[20].ReadOnly:=not TrouveMvt.Modif ;  // Cpt interdit
          Fliste.Columns.Items[19].ReadOnly:=Not TrouveMvt.Modif ;  // Contrepartie
          Fliste.Columns.Items[18].ReadOnly:=Not TrouveMvt.Modif ;  // Type de Contrepartie
          END Else
       If TrouveMvt.NATUREPIECE='S' then
          BEGIN
          NomCompte:=QEcr.FindField('S_SECTION').AsString+' '+QEcr.FindField('S_ABREGE').AsString ;
          END ;
       Caption:=NomCompte ;
       //Fliste.Columns.Items[Fliste.Columns.Items.indexof('G_GENERAL')].ReadOnly:=TrouveMvt.Modif ;
       END ;
   End ;
chercheInfo:=Ok ;
END ;

procedure TFMvtError.SEcrStateChange(Sender: TObject);
begin
Modifier:=True ;
end;

procedure TFMvtError.SEcrUpdateData(Sender: TObject);
begin
If Modifier then begin Modifier:=False ; If not OnSauve then bouge(nbcancel) ; end ;
end;

procedure TFMvtError.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(FListe,Nil,Caption,'') ;
end;

procedure TFMvtError.Button1Click(Sender: TObject);
begin
FicheGene(QEcr,'','', taConsult,0);

end;

procedure TFMvtError.BZOOMClick(Sender: TObject);
Var R : RLettR ;
begin
If TrouveMvt.indicetable=1 then
   BEGIN
   END Else
If TrouveMvt.indicetable=3 then  // Zoom sur le lettrage
   BEGIN
   FillChar(R,Sizeof(R),#0) ;
   R.General:=Qecr.FindField('E_GENERAL').AsString ; R.Auxiliaire:=Qecr.FindField('E_AUXILIAIRE').AsString ;
   R.CritDev:=Qecr.FindField('E_DEVISE').AsString ; R.GL:=NIL ; R.CritMvt:='' ; R.Appel:=tlMenu ;
   R.CodeLettre:=Qecr.FindField('E_LETTRAGE').AsString ; R.DeviseMvt:=R.CritDev ;
   R.LettrageDevise:=(Qecr.FindField('E_LETTRAGEDEV').AsString='X') ;
{$IFNDEF PGIIMMO}
   LettrageManuel(R,False,taConsult) ;
{$ENDIF}
   END Else
If TrouveMvt.indicetable=4 then // Zoom sur les comptes
   BEGIN
   if TrouveMvt.NATUREPIECE='G'then FicheGene(Nil,'',QEcr.FindField('G_GENERAL').AsString, taConsult,0) Else
   if TrouveMvt.NATUREPIECE='T'then FicheTiers(Nil,'',QEcr.FindField('T_AUXILIAIRE').AsString, taConsult,1) Else
   if TrouveMvt.NATUREPIECE='S'then FicheSection(Nil,QEcr.FindField('S_AXE').AsString,QEcr.FindField('S_SECTION').AsString, taConsult,0) Else

   if TrouveMvt.NATUREPIECE='J'then FicheJournal(Nil,'',QEcr.FindField('J_JOURNAL').AsString, taConsult,0) Else

   END ;
end;

procedure TFMvtError.BParamListeClick(Sender: TObject);
begin {ParamListe(QEcr.Liste,FListe,QEcr) ;}end;

Function TFMvtError.SiMvt : Boolean ;
BEGIN
Result:=TRUE ;
If TrouveMvt.NATUREPIECE='G' then
   Result:=ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+TrouveMvt.General+'" '+
                     ' AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+TrouveMvt.General+'"))'+
                     ' Or (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="'+TrouveMvt.General+'")))') else
If TrouveMvt.NATUREPIECE='T' then
   Result:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+TrouveMvt.Auxiliaire+'" '+
                     ' AND (EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+TrouveMvt.Auxiliaire+'"))') else
If TrouveMvt.NATUREPIECE='J' then
   Result:=ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="'+TrouveMvt.Journal+'" '+
                     ' AND ((EXISTS(SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL="'+TrouveMvt.Journal+'"))'+
                     ' Or (EXISTS(SELECT Y_JOURNAL FROM ANALYTIQ WHERE Y_JOURNAL="'+TrouveMvt.Journal+'")))') ;
END ;

procedure TFMvtError.FListeColEnter(Sender: TObject);
begin
IF TrouveMvt.IndiceTable<>4 then Exit ;
if Fliste.Selectedfield.ReadOnly=True then caption:=MsgLibel.Mess[3]+' '+NomCompte+' '+MsgLibel.Mess[4]
                                      else caption:=NomCompte ;
end;

procedure TFMvtError.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFMvtError.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=211 ;
end;

procedure TFMvtError.BRechercherClick(Sender: TObject);
begin FindFirst:=True; FindDialog.Execute ; end;

procedure TFMvtError.FindDialogFind(Sender: TObject);
begin Rechercher(FListe,FindDialog, FindFirst); end;

procedure TFMvtError.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
 
