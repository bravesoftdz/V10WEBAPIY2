unit SuprAuxi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, ComCtrls, HRichEdt, Grids, DBGrids,
  StdCtrls, Hctrls, ExtCtrls, Buttons, RapSuppr, hmsgbox, Mask, Hent1, Ent1,
  Hcompte, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE, ParamSoc,
{$IFNDEF GCGC}
  CPTiers_TOM,
{$ELSE GCGC}
  Fe_Main, AGLInit,
{$ENDIF GCGC}
   UtilSuprAuxi;

Procedure SuppressionCpteAuxi ( CliFou : String = '' ) ;

type
  TFSuprAuxi = class(TFMul)
    T_AUXILIAIRE: THCpteEdit;
    TT_AUXILIAIRE: TLabel;
    T_LIBELLE: TEdit;
    TT_LIBELLE: TLabel;
    T_NATUREAUXI: THValComboBox;
    TT_NATUREAUXI: TLabel;
    TT_COLLECTIF: TLabel;
    T_COLLECTIF: THCpteEdit;
    TT_DATECREATION: THLabel;
    T_DATECREATION: THCritMaskEdit;
    HLabel3: THLabel;
    T_DATECREATION_: THCritMaskEdit;
    TT_DATEDERNMVT: THLabel;
    T_DATEDERNMVT: THCritMaskEdit;
    TG_DATEDERNMVT2: THLabel;
    T_DATEDERNMVT_: THCritMaskEdit;
    T_DATEMODIF_: THCritMaskEdit;
    TG_DATEMODIFICATION2: THLabel;
    T_DATEMODIF: THCritMaskEdit;
    TT_DATEMODIFICATION: THLabel;
    T_MULTIDEVISE: TCheckBox;
    T_DEVISE: THValComboBox;
    TT_DEVISE: TLabel;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    BZauxi: TToolbarButton97;
    HTitres: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
    procedure T_NATUREAUXIChange(Sender: TObject);
  private
    Nblig     : Integer ;
    TDelAux  : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer : Boolean ;
    AuxiCode,GCCode  : String ;
    Function  Detruit ( St,RefGC : String): Byte ;
    Procedure Degage ;
//    Function  EstMouvemente(St : String) : Boolean  ;
  public
    CliFou : String;
  end;



implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure SuppressionCpteAuxi ( CliFou : String = '' ) ;
var FSuprAuxi : TFSuprAuxi ;
    PP       : THPanel ;
begin
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
FSuprAuxi:=TFSuprAuxi.Create(Application) ;
FSuprAuxi.FNomFiltre:='SUPPRAUXI' ; FSuprAuxi.Q.Manuel:=True ;
FSuprAuxi.Q.Liste:='SUPPRAUXITIERS' ; FSuprAuxi.CliFou:=CliFou ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FSuprAuxi.ShowModal ;
   finally
    FSuprAuxi.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FSuprAuxi,PP) ;
  FSuprAuxi.Show ;
  END ;
end ;

procedure TFSuprAuxi.FormShow(Sender: TObject);
begin
T_DATECREATION.Text:=StDate1900 ; T_DATECREATION_.Text:=StDate2099 ;
T_DATEMODIF.Text:=StDate1900 ; T_DATEMODIF_.Text:=StDate2099 ;
T_DATEDERNMVT.Text:=StDate1900 ; T_DATEDERNMVT_.Text:=StDate2099 ;
T_NATUREAUXI.ItemIndex:=0 ; T_DEVISE.ItemIndex:=0 ;
{$IFDEF GCGC}
if CliFou='CLI' then
   BEGIN
   T_NATUREAUXI.Value:='CLI' ; T_NATUREAUXI.Enabled:=False ;
   T_COLLECTIF.Enabled:=False ; TT_COLLECTIF.Enabled:=False ;
   Caption:=HTitres.Mess[0] ; UpdateCaption(Self) ;
   END else if CliFou='FOU' then
   BEGIN
   T_NATUREAUXI.Value:='FOU' ; T_NATUREAUXI.Enabled:=False ;
   T_COLLECTIF.Enabled:=False ; TT_COLLECTIF.Enabled:=False ;
   Caption:=HTitres.Mess[1] ; UpdateCaption(Self) ;
   END ;
{$ENDIF}
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFSuprAuxi.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X,Y : DelInfo ;
    Code,Lib,RefGC : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelAux.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN HM.execute(14,'','') ; Exit ; END ;
if HM.Execute(0,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('T_AUXILIAIRE').AsString ;
       RefGC:=Q.FindField('T_TIERS').AsString ;
       Lib:=Q.FindField('T_LIBELLE').AsString ;
       j:=Detruit(Code,RefGC) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[17] ;
          TDelAux.Add(X) ; Effacer:=True ;
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
   Code :=Q.FindField('T_AUXILIAIRE').AsString ;
   RefGC:=Q.FindField('T_TIERS').AsString ;
   Lib  :=Q.FindField('T_LIBELLE').AsString ;
   j:=Detruit(Code,RefGC) ;
   if j > 0 then
    begin
    if j in [1..13] then MsgDel.Execute(j-1,'','') else
     if j=14 then MsgDel.Execute(14,'','') else
      if j=17 then MsgDel.Execute(13,'','') else
       if j in [26..36] then MsgDel.Execute(13,'','');
    end else
    begin // DBR Fiche 1048 Debut
    Effacer := True;
    X:=DelInfo.Create;
    X.LeCod:=Code;
    X.LeLib:=Lib;
    X.LeMess:=HM.Mess[17] ;
    TDelAux.Add(X) ;
    NotEffacer := False;
    end; // DBR Fiche 1048 Fin
   END ;
if Effacer    then if HM.Execute(15,'','')=mrYes then RapportDeSuppression(TDelAux,1) ;
if NotEffacer then if HM.Execute(16,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;


Function TFSuprAuxi.Detruit ( St,RefGC : String):Byte ;
BEGIN
Result:=0 ;
if SupAEstMouvemente(St)  then BEGIN Result:=1 ;  Exit ; END ;
if SupAEstEcrGuide(St)    then BEGIN Result:=2 ;  Exit ; END ;
if SupAEstDansSociete(St) then BEGIN Result:=10 ; Exit ; END ;
if SupAEstDansSection(St) then BEGIN Result:=11 ; Exit ; END ;
if SupAEstDansUtilisat(St)then BEGIN Result:=13 ; Exit ; END ;
if SupAEstCpteCorresp(St) then BEGIN Result:=9 ;  Exit ; END ;
if SupAEstUnPayeur(St)    then BEGIN Result:=14 ; Exit ; END ;
{Gescom}
if SupAEstDansPiece(RefGC) then BEGIN Result:=26 ; Exit ; END ;
if SupAEstDansActivite(RefGC) then BEGIN Result:=28 ; Exit ; END ;
if SupAEstDansRessource(St) then BEGIN Result:=29 ; Exit ; END ;
if SupAEstDansAffaire(RefGC) then BEGIN Result:=30 ; Exit ; END ;
if SupAEstDansActions(St) then BEGIN Result:=31 ; Exit ; END ;
if SupAEstDansPersp(St) then BEGIN Result:=32 ; Exit ; END ;
if SupAEstDansCata(RefGC) then BEGIN Result:=34 ; Exit ; END ;
{Paie}
if SupAEstDansPaie(St) then BEGIN Result:=33 ; Exit ; END ;
if SupAEstDansMvtPaie(St) then BEGIN Result:=35 ; Exit ; END ;
if SupAEstDansProjets(RefGC) then BEGIN Result:=36 ; Exit ; END ;
AuxiCode:=St ; GCCode:=RefGC ;
if Transactions(Degage,5)<>oeOK then BEGIN MessageAlerte(HM.Mess[18]) ; Result:=17 ; Exit ; END ;
END ;

procedure TFSuprAuxi.Degage ;
Var DelTout : Boolean ;
BEGIN
DelTout:=True ;
if ExecuteSQL('DELETE FROM TIERS WHERE T_AUXILIAIRE="'+AuxiCode+'"')<>1 then
   BEGIN
   V_PGI.IoError:=oeUnknown ; Deltout:=False ;
   END ;
if DelTout then
   BEGIN
   ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+AuxiCode+'"') ;
   {Gescom}
   ExecuteSQL('DELETE FROM TIERSPIECE WHERE GTP_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSCOMPL WHERE YTC_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM ARTICLETIERS WHERE GAT_REFTIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TARIF WHERE GF_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+AuxiCode+'"') ;
   ExecuteSQL('DELETE FROM ADRESSES WHERE ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+GCCode+'"') ;
      //mcd 17/05/05 champ auxi supprimé ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'+GCCode+'"') ;
   ExecuteSQL('UPDATE ANNUAIRE SET ANN_TIERS="" WHERE ANN_TIERS="'+GCCode+'"') ;
   ExecuteSQL('DELETE FROM TIERSFRAIS WHERE GTF_TIERS="'+GCCode+'"');  { GPAO1_TIERSFRAIS }
   {Paie}
   END ;
END ;

procedure TFSuprAuxi.FormCreate(Sender: TObject);
begin
  inherited;
TDelAux:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprAuxi.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TDelAux.Clear ; TDelAux.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False,'',TRUE) ;
end;

procedure TFSuprAuxi.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
{$IFNDEF GCGC}
 FicheTiers(Q,'',Q.FindField('T_AUXILIAIRE').AsString,taConsult,0) ;
{$ELSE}
if CliFou='CLI' then AGLLanceFiche('GC','GCTIERS','',Q.FindField('T_AUXILIAIRE').AsString,ActionToString(taConsult)+';MONOFICHE;T_NATUREAUXI=CLI')
                else AGLLanceFiche('GC','GCTIERS','',Q.FindField('T_AUXILIAIRE').AsString,ActionToString(taConsult)+';MONOFICHE;T_NATUREAUXI=FOU') ;
{$ENDIF}
end;

procedure TFSuprAuxi.T_NATUREAUXIChange(Sender: TObject);
Var St : String ;
begin
  inherited;
St:=T_NATUREAUXI.Value ;
if St='' then T_AUXILIAIRE.ZoomTable:=tzTiers else
   if St='AUC' then T_AUXILIAIRE.ZoomTable:=tzTCrediteur else
      if St='AUD' then T_AUXILIAIRE.ZoomTable:=tzTDebiteur else
         if St='CLI' then T_AUXILIAIRE.ZoomTable:=tzTClient else
            if St='DIV' then T_AUXILIAIRE.ZoomTable:=tzTDivers else
               if St='FOU' then T_AUXILIAIRE.ZoomTable:=tzTFourn else
                  if St='SAL' then T_AUXILIAIRE.ZoomTable:=tzTsalarie ;
end;

{Function TFSuprAuxi.EstMouvemente(St : String) : Boolean  ;
BEGIN
Result:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" '+
                  'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+St+'" )') ;
END ;}

end.
