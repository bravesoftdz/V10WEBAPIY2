unit MulTiers;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, StdCtrls, Hcompte, Hctrls, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, //Tiers, //XMG 26/11/03
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1,
{$IFNDEF PGIIMMO}
{$IFNDEF CCMP}
{$IFNDEF IMP}
{$IFNDEF CCS3}
  MZSutil,
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
  HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil, HRichOLE,UtilPGI,
  ADODB, Mask ;

Procedure MultiCritereTiers(Comment : TActionFiche) ;

type
  TFMulTierC = class(TFMul)
    TT_LIBELLE: THLabel;
    T_LIBELLE: TEdit;
    T_EAN: TEdit;
    TT_EAN: THLabel;
    TT_DEVISE: THLabel;
    NePasVirer: TComboBox;
    HLabel1: THLabel;
    TT_AUXILIAIRE: THLabel;
    T_AUXILIAIRE: TEdit;
    T_NATUREAUXI: THValComboBox;
    TT_Secteur: THLabel;
    T_SECTEUR: THValComboBox;
    T_TARIFTIERS: THValComboBox;
    TT_TARIF: THLabel;
    TT_FACTURE: THLabel;
    TT_PAYEUR: THLabel;
    T_PAYEUR: THCpteEdit;
    T_FACTURE: THCpteEdit;
    T_LETTRABLE: TCheckBox;
    T_MULTIDEVISE: TCheckBox;
    HM: THMsgBox;
    T_REGIMETVA: THValComboBox;
    TT_REGIMETVA: THLabel;
    T_DEVISE: THValComboBox;
    T_COLLECTIF: THCpteEdit;
    TT_COLLECTIF: THLabel;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: TLabel;
    TT_TABLE1: TLabel;
    TT_TABLE2: TLabel;
    TT_TABLE3: TLabel;
    TT_TABLE4: TLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    TT_TABLE5: TLabel;
    T_TABLE5: THCpteEdit;
    TT_TABLE6: TLabel;
    T_TABLE6: THCpteEdit;
    TT_TABLE7: TLabel;
    T_TABLE7: THCpteEdit;
    TT_TABLE8: TLabel;
    T_TABLE8: THCpteEdit;
    TT_TABLE9: TLabel;
    T_TABLE9: THCpteEdit;
    T_ISPAYEUR: TCheckBox;
    XX_WHERE: TEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);  override;
    procedure BOuvrirClick(Sender: TObject);  override;
    procedure T_NATUREAUXIChange(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
  private
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses CPTiers_TOM,UFonctionsCBP;

Procedure MultiCritereTiers(Comment : TActionFiche) ;
var FMulTierC: TFMulTierC ;
    PP       : THPanel ;
begin
if Comment<>taConsult then if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulTierC:= TFMulTierC.Create(Application) ;
FMulTierC.TypeAction:=Comment ;
Case Comment Of
  taConsult      : begin
                   FMulTierC.Caption:=FMulTierC.HM.Mess[0];
                   FMulTierC.FNomFiltre:='MULVTIERS' ;
                   FMulTierC.Q.Liste:='MULVTIERS' ;
                   FMulTierC.HelpContext:=7139000 ;
                   end ;
  taModif        : begin
                   FMulTierC.Caption:=FMulTierC.HM.Mess[1];
                   FMulTierC.FNomFiltre:='MULMTIERS' ;
                   FMulTierC.Q.Liste:='MULMTIERS' ;
                   FMulTierC.HelpContext:=7145000 ;
                   end ;
  taModifEnSerie : begin
                   FMulTierC.Caption:=FMulTierC.HM.Mess[2];
                   FMulTierC.FNomFiltre:='MULMTIERS' ;
                   FMulTierC.Q.Liste:='MULMTIERS' ;
                   FMulTierC.HelpContext:=7148000 ;
                   end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulTierC.Caption:=FMulTierC.HM.Mess[10] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulTierC.ShowModal ;
    finally
     FMulTierC.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulTierC,PP) ;
   FMulTierC.Show ;
   END ;
end ;

procedure TFMulTierC.FormShow(Sender: TObject);
begin
if FCompte<>'' then T_AUXILIAIRE.Text:=FCompte ;
if FLibelle<>'' then T_Libelle.Text:=FLibelle ;
  inherited;
if TypeAction=taModifEnSerie then
   BEGIN
   FListe.MultiSelection := True ; BOuvrir.Hint:=HM.Mess[4] ;
   bSelectAll.Visible:=True ;
   END else
   BEGIN
   FListe.MultiSelection := False ;
   END;
// MODIF PACK AVANCE
if ((ExJaiLeDroitConcept(TConcept(ccAuxCreat),False)) and (TypeAction<>taConsult)) then BInsert.Visible:=True ;
{$IFDEF CCS3}
T_PAYEUR.Visible:=False ; TT_PAYEUR.Visible:=False ;
T_FACTURE.Visible:=False ; TT_FACTURE.Visible:=False ;
T_ISPAYEUR.Visible:=False ;
{$ENDIF}
end;

procedure TFMulTierC.FListeDblClick(Sender: TObject);
var Q1 : TQuery ;
    AA : TActionFiche ;
begin
if(Q.Eof) And (Q.Bof) then Exit ;
  inherited;
AA:=TypeAction ;
if Q.FindField('T_AUXILIAIRE')<>Nil then
   BEGIN
   if TypeAction<>taModifEnSerie then
      BEGIN
      if (TypeAction=taModif) then
        BEGIN
        Q1:=OpenSQL('SELECT T_FERME FROM TIERS WHERE T_AUXILIAIRE="'+Q.FindField('T_AUXILIAIRE').AsString+'"',True) ;
        if Q1.Fields[0].AsString='X' then HM.Execute(8,'','') ;
        Ferme(Q1) ;
        END ;
      if ((TypeAction=taModif) and (Not ExJaiLeDroitConcept(TConcept(ccAuxModif),False))) then AA:=taConsult ;
      (*{$IFDEF ESP}
      if V_PGI.MonoFiche then FicheTiers(Q.FindField('T_AUXILIAIRE').AsString,AA,0)
                         else FicheTiers(Q.FindField('T_AUXILIAIRE').AsString,AA,0) ;
      {$ELSE}*) //XMG 05/04/04
      if V_PGI.MonoFiche then FicheTiers(Nil,'',Q.FindField('T_AUXILIAIRE').AsString,AA,0)
                      else FicheTiers(Q,'',Q.FindField('T_AUXILIAIRE').AsString,AA,0) ;
      //{$ENDIF} //XMG 26/11/03 05/04/04
      if TypeAction<>taConsult then BChercheClick(Nil) ;
      END Else if TControl(Sender).Name='FListe' then
                  BEGIN
                  (*{$IFDEF ESP}
                  FicheTiers(Q.FindField('T_AUXILIAIRE').AsString,taModif,0) ;
                  {$ELSE}*) //XMG 05/04/04
                  FicheTiers(Q,'',Q.FindField('T_AUXILIAIRE').AsString,taModif,0) ;
                  //{$ENDIF} //XMG 26/11/03 05/04/04
                  Fliste.ClearSelected ;
                  END else
                  BEGIN
                  if (Fliste.NbSelected>0) or (FListe.AllSelected) then
                     BEGIN
{$IFNDEF PGIIMMO}
{$IFNDEF CCMP}
{$IFNDEF IMP}
{$IFNDEF CCS3}
                     ModifieEnSerie('TIERS','',FListe,Q) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
                     ChercheClick ;
                     END ;
                  END;
   END else HM.Execute(3,'','') ;
Screen.Cursor:=SyncrDefault ;
end;



Function TrouveAuxi(Var FCompte : string) : boolean ;
Var Cpte1 : String ;
    TAuxi : THTable ;
BEGIN
Result:=FALSE ;
TAuxi:=THTable.Create(Application) ;
TAuxi.DatabaseName:='SOC' ; TAuxi.TableName:='TIERS' ;
TAuxi.IndexName:='T_CLE1' ;
TAuxi.Open ;
TAuxi.FindNearest([FCompte]) ;
Cpte1:=TAuxi.FindField('T_AUXILIAIRE').AsString ;
if ((Copy(Cpte1,1,Length(FCompte))=FCompte) and (Not TAuxi.EOF)) then
   BEGIN
   TAuxi.Next ;
   if Copy(TAuxi.FindField('T_AUXILIAIRE').AsString,1,Length(FCompte))<>FCompte then
      BEGIN
      FCompte:=Cpte1 ; Result:=TRUE ;
      END ;
   END ;
TAuxi.Close ; TAuxi.Free ;
END ;


procedure TFMulTierC.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulTierC.T_NATUREAUXIChange(Sender: TObject);
Var NatTiers : String ;
begin
  inherited;
NatTiers:=T_NATUREAUXI.Value ;
if NatTiers='' then T_COLLECTIF.ZoomTable:=TzGCollectif else
   if NatTiers='CLI' then T_COLLECTIF.ZoomTable:=TzGCollClient else
      if NatTiers='FOU' then T_COLLECTIF.ZoomTable:=TzGCollFourn else
         if NatTiers='SAL' then T_COLLECTIF.ZoomTable:=TzGCollSalarie else
            if NatTiers='AUD' then T_COLLECTIF.ZoomTable:=TzGCollToutDebit else
               if NatTiers='AUC' then T_COLLECTIF.ZoomTable:=TzGCollToutCredit else
                  if NatTiers='DIV' then T_COLLECTIF.ZoomTable:=TzGCollDivers else
                     if NatTiers='NCP' then T_COLLECTIF.ZoomTable:=TzGCollectif ;
end;

procedure TFMulTierC.BParamListeClick(Sender: TObject);
begin
  inherited;
Fliste.ClearSelected ; 
end;

procedure TFMulTierC.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TT_TABLE','T_TABLE','T') ;
end;

procedure TFMulTierC.BinsertClick(Sender: TObject);
begin
  inherited;
// MODIF PACK AVANCE
(*{$IFDEF ESP}
  FicheTiers('',taCreatEnSerie,1) ;
{$ELSE}*) //XMG 05/04/04
  FicheTiers(Nil,'','',taCreatEnSerie,1) ;
//{$ENDIF} //XMG 26/11/03 05/04/04
end;

end.
