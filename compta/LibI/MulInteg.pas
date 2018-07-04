unit MulInteg;

// CA - 09/07/1999 - paramètres supplémentaires : intégration dans la compta et détail

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, DBTables, Hqry, ComCtrls, HRichEdt, HTB97,
  Grids, DBGrids, HDB, StdCtrls, ColMemo, Hctrls, ExtCtrls, Buttons, Mask,
  HEnt1, hmsgbox, ParamDat, Outils, ImEnt, Hstatus, UiUtil,
  HPanel, LookUp,UTOB,HRichOLE, ImOutGen;

type
  TFMulInteg = class(TFMul)
    HLabel1: THLabel;
    I_IMMO: TEdit;
    HLabel3: THLabel;
    I_LIBELLE: TEdit;
    Nature: THLabel;
    I_NATUREIMMO: THValComboBox;
    I_QUALIFIMMO: THValComboBox;
    HLabel2: THLabel;
    XX_WHERE: TEdit;
    TypeEcriture: THValComboBox;
    HLabel6: THLabel;
    HM: THMsgBox;
    HLabel11: THLabel;
    I_LIEUGEO: THValComboBox;
    HLabel5: THLabel;
    I_COMPTEIMMO: THCritMaskEdit;
    HLabel10: THLabel;
    I_COMPTELIE: THCritMaskEdit;
    bAbonnement: TToolbarButton97;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TI_TABLE0: TLabel;
    TI_TABLE5: TLabel;
    TI_TABLE6: TLabel;
    TI_TABLE1: TLabel;
    TI_TABLE2: TLabel;
    TI_TABLE3: TLabel;
    TI_TABLE4: TLabel;
    TI_TABLE7: TLabel;
    TI_TABLE8: TLabel;
    TI_TABLE9: TLabel;
    I_TABLE0: THCritMaskEdit;
    I_TABLE5: THCritMaskEdit;
    I_TABLE6: THCritMaskEdit;
    I_TABLE1: THCritMaskEdit;
    I_TABLE4: THCritMaskEdit;
    I_TABLE3: THCritMaskEdit;
    I_TABLE2: THCritMaskEdit;
    I_TABLE7: THCritMaskEdit;
    I_TABLE8: THCritMaskEdit;
    I_TABLE9: THCritMaskEdit;
    PzLibreS1: TTabSheet;
    Bevel7: TBevel;
    TT_TABLELIBREIMMO1: THLabel;
    TT_TABLELIBREIMMO2: THLabel;
    TT_TABLELIBREIMMO3: THLabel;
    TABLELIBRE3: THValComboBox;
    TABLELIBRE2: THValComboBox;
    TABLELIBRE1: THValComboBox;
    procedure bSelectAllClick(Sender: TObject);
    procedure FListeKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TypeEcritureChange(Sender: TObject);
    procedure OnCompteElipsisClick(Sender: TObject);
    procedure FListeCellClick(Column: TColumn);
    procedure FListeDblClick(Sender: TObject); override ;
    procedure bAbonnementClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override ;
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure TABLELIBRE1Change(Sender: TObject);
  private
    { Déclarations privées }
    procedure RecupereListeImmo(L: TStringList) ;
    {$IFDEF SERIE1}
    {$ELSE}
    procedure InitCommunContrat (Code,Libelle : string; OB : TOB);
    {$ENDIF}
  public
    { Déclarations publiques }
  end;

procedure AfficheMulIntegration(TypeInt: string='') ;
function CalculPeriodeEqAbo (Periode : string) : string;

implementation

uses ImContra, ImoGen,IntegEcr
{$IFDEF SERIE1}
 ;
{$ELSE}
   {$IFNDEF CCS3}
   ,Contabon
  {$ENDIF}
,ImCrGuid ;
{$ENDIF}

{$R *.DFM}

procedure AfficheMulIntegration(TypeInt: string='') ;
var
  FMulInteg: TFMulInteg;
  PP:THPanel;
begin
FMulInteg:=TFMulInteg.Create(Application) ;
FMulInteg.FNomFiltre:='MULVIMMOS' ;
FMulInteg.Q.Liste:='MULVIMMOS' ;
FMulInteg.CheckBoxStyle:=csCheckBox;
FMulInteg.TypeEcriture.Value:=TypeInt ;
PP:=FindInsidePanel;
if PP=nil then
begin
  try
    FMulInteg.ShowModal ;
  finally
    FMulInteg.Free ;
  end ;
end else
begin
  InitInside(FMulInteg,PP);
  FMulInteg.Show;
end;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulInteg.RecupereListeImmo(L: TStringList) ;
var i : integer;
begin
  if L= nil then exit ;
  if FListe.AllSelected then
  begin
    InitMove(Q.RecordCount,'');
    Q.First;
    while not Q.EOF do
    begin
      MoveCur(False);
      L.Add(Q.FindField('I_IMMO').AsString);
      Q.Next;
    end;
    FListe.AllSelected:=False;
  end else
  begin
    InitMove(FListe.NbSelected,'');
    for i:=0 to FListe.NbSelected-1 do
    begin
      MoveCur(False);
      FListe.GotoLeBookmark(i);
      L.Add(Q.FindField('I_IMMO').AsString );
    end;
    FListe.ClearSelected;
  end;
  FiniMove;
end;

procedure TFMulInteg.bSelectAllClick(Sender: TObject);
begin
  inherited;
  bAbonnement.Enabled :=  (fListe.nbSelected =1) and (TypeEcriture.value = 'ECH');
end;

procedure TFMulInteg.FListeKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  bAbonnement.Enabled :=  (fListe.nbSelected =1) and (TypeEcriture.value = 'ECH');  
end;

procedure TFMulInteg.FormShow(Sender: TObject);
begin
I_NATUREIMMO.ItemIndex:=0 ;
//I_NATUREIMMO.Value:='PRO' ;
{$IFDEF SERIE1}
if TypeEcriture.Value='' then TypeEcriture.Value := 'DOT'
                         else TypeEcriture.Enabled:=false ;
{$ELSE}
if TypeEcriture.Value='' then TypeEcriture.Value := 'DOT' ;
{$ENDIF}
TypeEcritureChange(nil) ;
{$IFDEF SERIE1}
if TypeEcriture.Value='DOT' then HelpContext:=541000 else HelpContext:=541100 ; 
{$ELSE}
HelpContext:=2410000 ;
{$ENDIF}

  inherited;
end;

procedure TFMulInteg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if IsInside(Self) then Action := caFree;
end;

procedure TFMulInteg.TypeEcritureChange(Sender: TObject);
begin
  inherited;
  if TypeEcriture.value = 'DOT' then
    XX_WHERE.Text:='(I_NATUREIMMO="PRO" OR I_NATUREIMMO="FI") AND I_ETAT<>"FER"'
  else if TypeEcriture.value = 'ECH' then
    XX_WHERE.Text:='(I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND I_ETAT<>"FER"'
  else XX_WHERE.Text:='I_ETAT<>"FER"';
end;

procedure TFMulInteg.OnCompteElipsisClick(Sender: TObject);
var stWhere : string;
begin
  inherited;
  if THCritMaskEdit(Sender).Name = 'I_COMPTEIMMO' then
  stWhere := '(G_GENERAL>="'+VHImmo^.CpteImmoInf+'" AND G_GENERAL<="'+VHImmo^.CpteImmoSup+'") OR '+
             '(G_GENERAL>="'+VHImmo^.CpteFinInf+'" AND G_GENERAL<="'+VHImmo^.CpteFinSup+'")'
  else if THCritMaskEdit(Sender).Name = 'I_COMPTELIE' then
    stWhere := '(G_GENERAL>="'+VHImmo^.CpteCBInf+'" AND G_GENERAL<="'+VHImmo^.CpteCBSup+'") OR '+
               '(G_GENERAL>="'+VHImmo^.CpteLocInf+'" AND G_GENERAL<="'+VHImmo^.CpteLocSup+'")';
  LookupList(TControl(Sender),'','GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
end;

procedure TFMulInteg.FListeCellClick(Column: TColumn);
begin
  inherited;
  bAbonnement.Enabled :=  (fListe.nbSelected =1) and (TypeEcriture.value = 'ECH');
end;

procedure TFMulInteg.FListeDblClick(Sender: TObject);
begin
  inherited;
  bAbonnement.Enabled :=  (fListe.nbSelected =1) and (TypeEcriture.value = 'ECH');
  FicheImmobilisation(Q,Q.FindField('I_IMMO').AsString,taConsult,'') ;
end;

procedure TFMulInteg.bAbonnementClick(Sender: TObject);
{$IFDEF SERIE1}
begin
{$ELSE}
{$IFDEF CCS3}
BEGIN
{$ELSE}
var OB, OBAbo : TOB;
    QImmo, QContrat : TQuery;
    Contrat : TImContrat;
    stCodeAbo, stPeriode : string;
    i : integer;
    ARecord : PTranche;
begin
  inherited;
  stCodeAbo := Q.FindField('I_NATUREIMMO').AsString + Q.FindField('I_IMMO').AsString + '01';
  QContrat := OpenSQL ('SELECT * FROM CONTABON WHERE CB_CONTRAT="'+stCodeAbo+'"',True);
  if not QContrat.Eof then ParamAbonnement(True,stCodeAbo,taConsult)
  else
  begin
    QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+Q.FindField('I_IMMO').AsString+'"',True);
    stPeriode := QImmo.FindField('I_PERIODICITE').AsString;
    Contrat := TImContrat.Create;
    Contrat.Charge (QImmo);
    Contrat.ChargeTableEcheance;
    Ferme (QImmo);
    OB := TOB.Create ('',nil,-1);
    Contrat.ConvertEcheanceIntoTranche;
    if Contrat.sTypeLoyer = 'LCO' then
    begin // Cas loyer constant
      OBAbo := TOB.Create ('CONTABON',OB,-1);
      InitCommunContrat (stCodeAbo,Q.FindField('I_LIBELLE').AsString,OBAbo);
      if Contrat.ListeEcheances.Count = 2 then ARecord := Contrat.ListeTranches.Items[1]
      else ARecord := Contrat.ListeTranches.Items[1];
      OBAbo.PutValue('CB_DATECONTRAT',ARecord^.DateDebut);
      OBAbo.PutValue('CB_SEPAREPAR',CalculPeriodeEqAbo (stPeriode));
      OBAbo.PutValue('CB_NBREPETITION',ARecord^.nEcheance);
      CreationGuideEcheance (OBAbo);
    end
    else
    begin  // Cas  loyer variable
      for i := 1 to Contrat.ListeTranches.Count do
      begin
        OBAbo := TOB.Create ('CONTABON',OB,-1);
        stCodeAbo := Q.FindField('I_NATUREIMMO').AsString + Contrat.sCode + Format('%.2d',[i]);
        ARecord := Contrat.ListeTranches.Items[i-1];
        InitCommunContrat (stCodeAbo,Q.FindField('I_LIBELLE').AsString,OBAbo);
        OBAbo.PutValue('CB_DATECONTRAT',ARecord^.DateDebut);
        OBAbo.PutValue('CB_SEPAREPAR',CalculPeriodeEqAbo (stPeriode));
        OBAbo.PutValue('CB_NBREPETITION',ARecord^.nEcheance);
        CreationGuideEcheance (OBAbo);
      end;
    end;
    OB.InsertDB(nil);
    Contrat.Free;
    OB.Free;
  end;
  Ferme (QContrat);
{$ENDIF}
{$ENDIF}
end;

{$IFDEF SERIE1}
{$ELSE}
procedure TFMulInteg.InitCommunContrat (Code,Libelle : string; OB : TOB);
begin
  OB.PutValue('CB_CONTRAT',Code);
  OB.PutValue('CB_COMPTABLE','X');
  OB.PutValue('CB_LIBELLE',Libelle);
  OB.PutValue('CB_ARRONDI','PAS');
  OB.PutValue('CB_RECONDUCTION','SUP');
  OB.PutValue('CB_DEJAGENERE',0);
  OB.PutValue('CB_DATEDERNGENERE',iDate1900);
  OB.PutValue('CB_GUIDE','');
  OB.PutValue('CB_DATECREATION',Date);
  OB.PutValue('CB_DATEMODIF',Date);
  OB.PutValue('CB_UTILISATEUR',V_PGI.User);
  OB.PutValue('CB_SOCIETE',V_PGI.CodeSociete);
end;
{$ENDIF}
function CalculPeriodeEqAbo (Periode : string) : string;
begin
  if Periode = 'MEN' then Result := '1M'
  else if Periode = 'TRI' then Result := '3M'
  else if Periode = 'SEM' then Result := '6M'
  else if Periode = 'ANN' then Result := '12M'
  else Result := '1M';
end;

procedure TFMulInteg.BOuvrirClick(Sender: TObject);
var ListeImmo : TStringList;
    Nature : string;
begin
  inherited;
  ListeImmo:=TStringList.Create ;
  try
    RecupereListeImmo(ListeImmo) ;
    if ListeImmo.Count <> 0 then
    begin
      Q.First;
      Nature := Q.FindField('I_NATUREIMMO').AsString;
      if (Nature = 'PRO') or (Nature = 'FI') then
        IntegrationEcritures (toDotation,ListeImmo,TRUE,FALSE)
      else if (Nature = 'CB') or (Nature = 'LOC') then
        IntegrationEcritures (toEcheance,ListeImmo,TRUE,FALSE);
    end;
    FListe.ClearSelected;
  finally
    ListeImmo.Free ;
  end ;
end;

procedure TFMulInteg.HMTradBeforeTraduc(Sender: TObject);
var Okok: boolean ;
begin
  inherited;
  {$IFDEF SERIE1}
  ImLibellesTableLibre(PzLibreS1,'TT_TABLELIBREIMMO','','I') ;
  Okok:=false ;
  {$ELSE}
  ImLibellesTableLibre(PzLibre,'TI_TABLE','I_TABLE','I') ;
  Okok:=true ;
  {$ENDIF}
  PzLibre.TabVisible:=Okok ;
  PzLibreS1.TabVisible:=not Okok ;
end;

procedure TFMulInteg.TABLELIBRE1Change(Sender: TObject);
begin
  inherited;
  if Sender=TABLELIBRE1 then I_TABLE0.Text:=TABLELIBRE1.Value
  else if Sender=TABLELIBRE2 then I_TABLE1.Text:=TABLELIBRE2.Value
  else if Sender=TABLELIBRE3 then I_TABLE2.Text:=TABLELIBRE3.Value
end;

end.

