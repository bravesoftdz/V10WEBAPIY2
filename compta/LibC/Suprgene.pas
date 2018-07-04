{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... :   /  /    
Description .. : Remplace en eAGL par UTOFMULPARAMGEN.PAS
Mots clefs ... : 
*****************************************************************}
unit Suprgene;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, Hqry, Grids,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1,
  CPGeneraux_TOM,
  HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  ADODB ;

Procedure SuppressionCpteGene ;

type
  TFSuprgene =  class(TFMul)
    TG_GENERAL: THLabel;
    TG_ABREGE: THLabel;
    G_LIBELLE: TEdit;
    TG_NATUREGENE: THLabel;
    G_NATUREGENE: THValComboBox;
    G_COLLECTIF: TCheckBox;
    G_POINTABLE: TCheckBox;
    G_LETTRABLE: TCheckBox;
    TG_SENS: THLabel;
    G_SENS: THValComboBox;
    TG_DATEMODIFICATION: THLabel;
    G_DATEMODIF: THCritMaskEdit;
    TG_DATEMODIFICATION2: THLabel;
    G_DATEMODIF_: THCritMaskEdit;
    G_DATEDERNMVT_: THCritMaskEdit;
    TG_DATEDERNMVT2: THLabel;
    G_DATEDERNMVT: THCritMaskEdit;
    TG_DATEDERNMVT: THLabel;
    G_DATECREATION_: THCritMaskEdit;
    HLabel3: THLabel;
    G_DATECREATION: THCritMaskEdit;
    HLabel2: THLabel;
    BZgene: TToolbarButton97;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    G_GENERAL: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
    procedure G_NATUREGENEChange(Sender: TObject);
  private    { Déclarations privées }
    Nblig     : Integer ;
    TDelGene  : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer : Boolean ;
    GeneCode  : String ;
    Function  Detruit(St : String): Byte ;
    Function  EstMouvemente(St : String) : Boolean ;
    Procedure Degage ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses UtilPgi ;
Procedure SuppressionCpteGene ;
var FSuprgene : TFSuprgene;
    PP       : THPanel ;
begin
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
FSuprgene:=TFSuprGene.Create(Application) ;
FSuprgene.FNomFiltre:='SUPPRGENE' ; FSuprgene.Q.Manuel:=True ; FSuprgene.Q.Liste:='SUPPRGENE' ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FSuprgene.ShowModal ;
   finally
    FSuprgene.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FSuprgene,PP) ;
  FSuprgene.Show ;
  END ;
end ;

procedure TFSuprgene.FormShow(Sender: TObject);
begin
G_DATECREATION.Text:=StDate1900 ; G_DATECREATION_.Text:=StDate2099 ;
G_DATEMODIF.Text:=StDate1900 ; G_DATEMODIF_.Text:=StDate2099 ;
G_DATEDERNMVT.Text:=StDate1900 ; G_DATEDERNMVT_.Text:=StDate2099 ;
G_NATUREGENE.ItemIndex:=0 ; G_SENS.ItemIndex:=0 ;
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFSuprgene.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X,Y : DelInfo ;
    Code,Lib : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelGene.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN HM.execute(13,'','') ; Exit ; END ;
if HM.Execute(0,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('G_GENERAL').AsString ;
       Lib:=Q.FindField('G_LIBELLE').AsString ;
       j:=Detruit(Code) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[16] ;
          TDelGene.Add(X) ; Effacer:=True ;
          END else
          BEGIN
          Y:=DelInfo.Create ; Y.LeCod:=Code ; Y.LeLib:=Lib ;
          Y.LeMess:=HM.Mess[j] ;
          TNotDel.Add(Y) ;  NotEffacer:=True ;
          END
      END ;
   END else
   BEGIN
   Fliste.GotoLeBookMark(0) ;
   Code:=Q.FindField('G_GENERAL').AsString ; j:=Detruit(Code) ;
   if j in [1..12] then MsgDel.Execute(j-1,'','') else
      if j=7 then MsgDel.Execute(12,'','')
   END ;
if Effacer    then if HM.Execute(14,'','')=mrYes then RapportDeSuppression(TDelGene,1) ;
if NotEffacer then if HM.Execute(15,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

Function TFSuprgene.EstMouvemente(St : String):Boolean ;
BEGIN
Result:=ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+St+'" '+
                  ' AND ((EXISTS(SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+St+'"))'+
                  ' Or (EXISTS(SELECT Y_GENERAL FROM ANALYTIQ WHERE Y_GENERAL="'+St+'")))') ;
END ;

procedure TFSuprgene.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM GENERAUX WHERE G_GENERAL="'+GeneCode+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

Function TFSuprgene.Detruit(St : String):Byte ;
BEGIN
Result:=0 ;
if EstMouvemente(St)   then BEGIN Result:=1 ; Exit ; END ;
if EstAnalytiqPure(St) then BEGIN Result:=3 ; Exit ; END ;
if EstCpteAxe(St)      then BEGIN Result:=4 ; Exit ; END ;
//if EstBanquaire(St)    then BEGIN Result:=5 ; Exit ; END ;
if EstCpteCorresp(St)  then BEGIN Result:=6 ; Exit ; END ;
if EstCpteDevise(St)   then BEGIN Result:=7 ; Exit ; END ;
//if EstCpteDimmo        then BEGIN Result:=8 ; Exit ; END ;
// Table IMMO Non encore Créée
if EstCpteJournal(St)  then BEGIN Result:=9 ; Exit ; END ;
if EstCpteModepaie(St) then BEGIN Result:=10; Exit ; END ;
if EstCpteSociete(St)  then BEGIN Result:=11; Exit ; END ;
if EstCpteTva(St)      then BEGIN Result:=12; Exit ; END ;
GeneCode:=St ;
if Transactions(Degage,5)<>oeOK then
   BEGIN
   MessageAlerte(HM.Mess[17]) ; Result:=17 ; Exit ;
   END else
   BEGIN
   ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="'+GeneCode+'" And V_NATURE Like"GE%"') ;
   ExecuteSql('DELETE From BANQUECP Where BQ_GENERAL="'+GeneCode+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"') ; // 19/10/2006 YMO Multisociétés
   END ;
END ;

procedure TFSuprgene.FormCreate(Sender: TObject);
begin
  inherited;
TDelgene:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprgene.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TDelGene.Clear ; TDelGene.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False,'',TRUE) ;
end;

procedure TFSuprgene.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheGene(Q,'',Q.FindField('G_GENERAL').AsString,taConsult,0) ;
end;

procedure TFSuprgene.G_NATUREGENEChange(Sender: TObject);
Var St : String ;
begin
  inherited;
St:=G_NATUREGENE.Value ;
if St='' then G_GENERAL.ZoomTable:=tzGeneral else
   if St='BQE' then G_GENERAL.ZoomTable:=tzGbanque else
      if St='CAI' then G_GENERAL.ZoomTable:=tzGcaisse else
         if St='CHA' then G_GENERAL.ZoomTable:=tzGcharge else
            if St='COC' then G_GENERAL.ZoomTable:=TzGCollClient else
               if St='COD' then G_GENERAL.ZoomTable:=tzGCollDivers else
                  if St='COS' then G_GENERAL.ZoomTable:=tzGCollSalarie else
                     if St='COF' then G_GENERAL.ZoomTable:=tzGCollFourn else
                        if St='DIV' then G_GENERAL.ZoomTable:=tzGDivers else
                           if St='EXT' then G_GENERAL.ZoomTable:=tzGextra else
                              if St='IMO' then G_GENERAL.ZoomTable:=tzGimmo else
                                 if St='PRO' then G_GENERAL.ZoomTable:=tzGproduit else
                                    if St='TIC' then G_GENERAL.ZoomTable:=tzGTIC else
                                       if St='TID' then G_GENERAL.ZoomTable:=tzGTID ;
end;

end.
