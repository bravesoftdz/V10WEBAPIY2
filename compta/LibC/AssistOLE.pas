{***********UNITE*************************************************
Auteur  ...... : ???
Créé le ...... : 01/01/1900
Modifié le ... : 17/08/2005
Description .. :
Suite ........ : FQ 13827 ( SBO 17/08/2005 ) :
Suite ........ :  - Ajout de la gestion des IFRS
Mots clefs ... :
*****************************************************************}
unit AssistOLE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, DB, DBTables, StdCtrls, Hctrls, Grids, DBGrids,
{$IFNDEF EAGLCLIENT}
  HDB,
{$ENDIF}
  ComCtrls,
  Buttons, ExtCtrls, hmsgbox, HEnt1, HSysMenu, HTB97, Hqry, ColMemo,
  UTOB, HPanel;

function LanceAssistOlE(Formule : String) : String ;

Type tav = (avNone,avEcr,avAna,avCpt) ;

type
  TFAssistOLE = class(TFAssist)
    ChoixRec: TTabSheet;
    Periode: TTabSheet;
    MsgEF: THMsgBox;
    FindCode: TEdit;
    FindLib: TEdit;
    LExercice: THLabel;
    cExercice: THValComboBox;
    HLabel7: THLabel;
    cPeriode: THValComboBox;
    cDetailPeriode: THValComboBox;
    NumPeriode: THNumEdit;
    lDetailPeriode: THLabel;
    Criteres: TTabSheet;
    HLabel5: THLabel;
    cNatureEcr: THValComboBox;
    HLabel3: THLabel;
    cIntegAN: THValComboBox;
    lRub: THLabel;
    TitleLine: TGroupBox;
    CheckExoOK: TCheckBox;
    cNumPeriode: THValComboBox;
    Etablissement: TTabSheet;
    Label1: TLabel;
    cEtablissement: THValComboBox;
    Devise: TTabSheet;
    Label4: TLabel;
    cDevise: THValComboBox;
    FlagExtEcr: TCheckBox;
    RefCelEcr: TEdit;
    FlagExtPer: TCheckBox;
    RefCelPer: TEdit;
    FlagExtRec: TCheckBox;
    RefCelRec: TEdit;
    FlagExtEtab: TCheckBox;
    RefCelEtab: TEdit;
    FlagExtDev: TCheckBox;
    RefCelDev: TEdit;
    ChoixSource: TTabSheet;
    rSource: TRadioGroup;
    ChoixFormule: TTabSheet;
    Label2: TLabel;
    Label3: TLabel;
    rCumul: TRadioButton;
    rChamp: TRadioButton;
    ChoixChamp: TTabSheet;
    HLabel1: THLabel;
    FlagExtField: TCheckBox;
    RefCelField: TEdit;
    lSource: TLabel;
    Label5: TLabel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    Resultat: TTabSheet;
    HLabel10: THLabel;
    rConstante: TRadioButton;
    cLibelle: TCheckBox;
    bVal: TSpeedButton;
    bLib: TSpeedButton;
    MsgF: THMsgBox;
    Panel1: TPanel;
    HLabel8: THLabel;
    l1: THLabel;
    l2: THLabel;
    l3: THLabel;
    l4: THLabel;
    l5: THLabel;
    l6: THLabel;
    pFormule: THLabel;
    r1: THLabel;
    r2: THLabel;
    r3: THLabel;
    r4: THLabel;
    r5: THLabel;
    r6: THLabel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    Bevel8: TBevel;
    Bevel9: TBevel;
    Bevel1: TBevel;
    Revision: TRadioGroup;
    rBudget: TRadioGroup;
    lBudget: TLabel;
    gAxe: TGroupBox;
    lAxe: THLabel;
    cAxe: THValComboBox;
    TabSheet1: TTabSheet;
    Label6: TLabel;
    HValComboBox1: THValComboBox;
    CheckBox1: TCheckBox;
    Edit1: TEdit;
    Budget: TTabSheet;
    Label7: TLabel;
    cBudget: THValComboBox;
    FlagExtBud: TCheckBox;
    RefCelBud: TEdit;
    CptVariant: TTabSheet;
    FlagExtVariant: TCheckBox;
    RefCellVariant: TEdit;
    Label8: TLabel;
    cCptVariant: TEdit;
    NumPeriode1: THNumEdit;
    LNumperiode1: THLabel;
    cAnnee: THValComboBox;
    Label9: TLabel;
    GSens: TGroupBox;
    LSens: THLabel;
    CSens: THValComboBox;
    CMonnaie: TRadioGroup;
    ChoixRec2: TTabSheet;
    FindCode2: TEdit;
    FindLib2: TEdit;
    FlagExtRec2: TCheckBox;
    RefCelRec2: TEdit;
    GC2: TCheckBox;
    RSQL: TRadioButton;
    PSQL: TTabSheet;
    GroupBox2: TGroupBox;
    FlagExtSQL: TCheckBox;
    RefCelSQL: TEdit;
    Z_SQL: THSQLMemo;
    PAvances: TTabSheet;
    bEffaceAvance: TToolbarButton97;
    Z_C3: THValComboBox;
    Z_C2: THValComboBox;
    Z_C1: THValComboBox;
    ZO1: THValComboBox;
    ZO2: THValComboBox;
    ZO3: THValComboBox;
    ZV3: TEdit;
    ZV2: TEdit;
    ZV1: TEdit;
    ZG1: TComboBox;
    ZG2: TComboBox;
    Z_C4: THValComboBox;
    Z_C5: THValComboBox;
    Z_C6: THValComboBox;
    ZO6: THValComboBox;
    ZO5: THValComboBox;
    ZO4: THValComboBox;
    ZV4: TEdit;
    ZV5: TEdit;
    ZV6: TEdit;
    ZG4: TComboBox;
    ZG5: TComboBox;
    ZG3: TComboBox;
    CBLib: TCheckBox;
    rEcr: TRadioGroup;
    ChoixBalSit: TTabSheet;
    Label10: TLabel;
    CBalSit: THValComboBox;
    FlagExtBalSit: TCheckBox;
    RefCelBalSit: TEdit;
    GBalSit: TGroupBox;
    CCBalSit: TCheckBox;
    HLabel6: THLabel;
    Bevel2: TBevel;
    r7: THLabel;
    Regroupement: TLabel;
    MULTIDOSSIER: THValComboBox;
    AvecIFRS: TCheckBox;
    THGRID: THGrid;
    GridRub2: THGrid;
    GridFields: THGrid;
    procedure cPeriodeChange(Sender: TObject);
    procedure cNatureEcrEnter(Sender: TObject);
    procedure cIntegANEnter(Sender: TObject);
    procedure cExerciceEnter(Sender: TObject);
    procedure cPeriodeEnter(Sender: TObject);
    procedure NumPeriodeEnter(Sender: TObject);
    procedure cDetailPeriodeEnter(Sender: TObject);
    procedure cEtablissementEnter(Sender: TObject);
    procedure cDeviseEnter(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure GridRubEnter(Sender: TObject);
    procedure FindCodeEnter(Sender: TObject);
    procedure FindLibEnter(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckExoOKClick(Sender: TObject);
    procedure cNumPeriodeClick(Sender: TObject);
    function  FirstPage : TTabSheet ; Override ;
    function  PreviousPage : TTabSheet ; Override ;
    function  NextPage : TTabSheet ; Override ;
//    function  PageCount : Integer ; Override ;
    function  PageNumber : Integer ; Override ;
    procedure FlagExtClick(Sender: TObject);
    procedure RefCelEnter(Sender: TObject);
    procedure FlagExtEnter(Sender: TObject);
    procedure RefCelRecChange(Sender: TObject);
    procedure rSourceClick(Sender: TObject);
    procedure PChange(Sender: TObject); Override ;
    procedure rSourceEnter(Sender: TObject);
    procedure cAxeEnter(Sender: TObject);
    procedure GridRubDblClick(Sender: TObject);
    procedure rCumulClick(Sender: TObject);
    procedure rChampClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GridFieldsDblClick(Sender: TObject);
    procedure rConstanteClick(Sender: TObject);
    procedure NumPeriodeExit(Sender: TObject);
    procedure bValClick(Sender: TObject);
    procedure bLibClick(Sender: TObject);
    procedure rBudgetClick(Sender: TObject);
    procedure cAxeChange(Sender: TObject);
    procedure cCptVariantEnter(Sender: TObject);
    procedure NumPeriodeChange(Sender: TObject);
    procedure NumPeriode1Exit(Sender: TObject);
    procedure FindCode2Change(Sender: TObject);
    procedure FindLib2Change(Sender: TObject);
    procedure RSQLClick(Sender: TObject);
    procedure CBLibClick(Sender: TObject);
    procedure bEffaceAvanceClick(Sender: TObject);
    procedure rEcrClick(Sender: TObject);
    procedure CCBalSitClick(Sender: TObject);
    procedure GridFieldsClick(Sender: TObject);
    procedure THGRIDClick(Sender: TObject);
    procedure GridRub2Click(Sender: TObject);
    procedure FindCodeChange(Sender: TObject);
  private  { Déclarations privées }
    sFormule,InitialRec,InitialRec2,InitialField,InitialSQL : String ;
    InitialFormula : String ;
    LoadingFormula : Boolean ;
    SauveAxe : integer ;
    AxeCourant : string ;
    LeCptVariant : String ;
    function  GetMonth(i : Integer) : String ;
    function  GetTwoMonth(i1,i2 : Integer) : String ;
    function  GetNatureEcr(s : String) : String ;
    procedure ChangeSourceItems(n : Integer) ;
    procedure SetInitialValues ;
    procedure SetInitialTable(s : String) ;
    procedure SetInitialRec(s : String) ;
    procedure SetInitialRec2(s : String) ;
    procedure SetInitialBalSit(s : String) ;
    Procedure DecodeWhereSup(S : String) ;
    procedure SetInitialWhereSup(s : String) ;
    procedure SetInitialField(s : String) ;
    procedure SetInitialEcr(s : String) ;
    procedure SetInitialEtab(s : String) ;
    procedure SetInitialSQL(s : String) ;
    procedure SetInitialDev(s : String) ;
    procedure SetInitialCptVariant(s : String) ;
    procedure SetInitialDate(s : String ; PourConstante : Boolean) ;
    Function  WriteFormule : Boolean ;
    procedure ExpliqueFormule ;
    function  IsSociete : boolean ;
    function  IsBudgetaire : boolean ;
    function  IsSection : boolean ;
    function  IsSection2 : boolean ;
    function  IsCptGeneral : boolean ;
    function  IsBudJal : boolean ;
    function  IsBalSit : boolean ;
    function  IsRubrique : boolean ;
    function  IsBudgete : boolean ;
    function  WhereBudjal : string ;
//    function  AvecCptVariant : boolean ;
    procedure ChargeAvances(OnLib : Boolean) ;
    Function  FaitSQLSupp : String ;
    Function  OkAvances(Recharge : Boolean ; ForceIndex : Integer = -1) : tav ;
    procedure ChangementCode(var Grille : THGrid; FindCod, FindLibelle : string);
  public   { Déclarations publiques }
  end;

implementation

{$R *.DFM}


uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  Ent1;

function LanceAssistOLE(formule : String) : string ;
Var X : TFAssistOLE ;
    Qzi : Integer ;
begin
Result:='' ;
SourisSablier ;
X:=TFAssistOLE.Create(Application) ;
 Try
  X.Caption := X.MsgEF.Mess[0] ;
  X.InitialFormula := Formule ;
  QZI:=X.ShowModal ;
  if Qzi=mrOK then Result:=X.sFormule ;
 Finally
  X.Free ;
 end ;
Screen.Cursor:= crDefault ;
SourisNormale ;
end ;

function TFAssistOLE.IsSociete : boolean ;
begin
result := (rSource.ItemIndex=9) ;
end ;

function TFAssistOLE.IsBudgetaire : boolean ;
begin
result := (rCumul.Checked) and ((rSource.ItemIndex=1) or (rSource.ItemIndex=5) or (rSource.ItemIndex=6) Or (rSource.ItemIndex=12)) ;
end ;

function TFAssistOLE.IsSection : boolean ;
begin
result := (rSource.ItemIndex=4) or (rSource.ItemIndex=6) ;
end ;

function TFAssistOLE.IsSection2 : boolean ;
begin
result := (rSource.ItemIndex  in [11,12]);
end ;


function TFAssistOLE.IsCptGeneral : boolean ;
begin
result := (rSource.ItemIndex=2) ;
end ;

function TFAssistOLE.IsBudJal : boolean ;
begin
result := False ;
if rCumul.Checked then result := (rSource.ItemIndex=1) or ((rBudget.Visible) and (rBudget.ItemIndex=0)) ;
if rChamp.Checked then result := (rSource.ItemIndex=1) or (rSource.ItemIndex=5) or (rSource.ItemIndex=6) ;
end ;

function TFAssistOLE.IsBalSit : boolean ;
begin
Result:=FALSE ;
If rCumul.Checked Then result:=(CCBalSit.Checked)
End ;

function TFAssistOLE.IsRubrique : boolean ;
begin
result := False ;
if rCumul.Checked then result := (rSource.ItemIndex=0) or (rSource.ItemIndex=1) ;
end ;

{
function TFAssistEF.AvecCptVariant : boolean ;
begin
result := False ;
if trim(cCptVariant.Text)<>'' then result := (rSource.ItemIndex=0) or (rSource.ItemIndex=1) ;
end ;
}

function TFAssistOLE.IsBudgete : boolean ;
begin
result := (rBudget.Visible) and (rBudget.ItemIndex=0) ;
end ;

function TFAssistOLE.WhereBudJal : string ;
begin
if (rCumul.Checked) and (cBudget.Visible) and (cBudget.Value<>'') then result := 'RB_BUDJAL="' + cBudget.Value + '" AND ' else result := '' ;
end ;

function TFAssistOLE.GetMonth(i : Integer) : String ;
begin
result := MsgEF.Mess[12+i] ;
end ;

function TFAssistOLE.GetTwoMonth(i1,i2 : Integer) : String ;
begin
result := MsgEF.Mess[25] + ' ' + Lowercase(GetMonth(i1)) + ' ' + MsgEF.Mess[26] + ' ' + Lowercase(GetMonth(i2)) ;
end ;

function TFAssistOLE.GetNatureEcr(s : String) : String ;
var sr : String ;
begin
if s = 'TOU'    then sr := 'NSUP' ;
if s = 'NOR' then sr := 'N' ;
if s = 'NSS' then sr := 'NSU' ;
if s = 'SSI' then sr := 'SU' ;
if s = 'PRE' then sr := 'P' ;
if Revision.ItemIndex=0 then sr := sr + 'R' ;
if AvecIFRS.Checked then sr := sr + 'I' ; // FQ 13827 : SBO 17/082005
result := sr ;
end ;

function TFAssistOLE.FirstPage : TtabSheet  ;
begin
result := ChoixFormule ;
end ;

function TFAssistOLE.PageNumber : Integer ;
begin
result := 0 ;
if rCumul.Checked then
   begin
   if GetPageName = 'ChoixFormule' then result := 1 ;
   if GetPageName = 'ChoixSource' then result := 2 ;
   if GetPageName = 'Budget' then result := 3 ;
   if GetPageName = 'CptVariant' then result := 4 ;
   if GetPageName = 'ChoixRec' then result := 5 ;
   if GetPageName = 'ChoixRec2' then result := 6 ;
   if GetPageName = 'Periode' then result := 7 ;
   if GetPageName = 'Criteres' then result := 8 ;
   if GetPageName = 'Etablissement' then result := 9 ;
   if GetPageName = 'Devise' then result := 10 ;
   if GetPageName = 'ChoixBalSit' then result := 11 ;
   if GetPageName = 'PAvances' then result := 12 ;
   if GetPageName = 'Resultat' then result := 13 ;
   end ;
if rChamp.Checked then
   begin
   if GetPageName = 'ChoixFormule' then result := 1 ;
   if GetPageName = 'ChoixSource' then result := 2 ;
   if GetPageName = 'Budget' then result := 3 ;
   if GetPageName = 'ChoixRec' then result := 4 ;
   if GetPageName = 'ChoixChamp' then result := 5 ;
   if GetPageName = 'Resultat' then result := 6 ;
   end ;
if rConstante.Checked then
   begin
   if GetPageName = 'ChoixFormule' then result := 1 ;
   if GetPageName = 'ChoixRec' then result := 2 ;
   if GetPageName = 'Etablissement' then result := 3 ;
   if GetPageName = 'Periode' then result := 4 ;
   if GetPageName = 'Resultat' then result := 5 ;
   end ;
if rSQL.Checked then
   begin
   if GetPageName = 'ChoixFormule' then result := 1 ;
   if GetPageName = 'PSQL' then result := 2 ;
   if GetPageName = 'Resultat' then result := 3 ;
   end ;
end ;

function TFAssistOLE.PreviousPage : TTabSheet ;
begin
result := nil ;
if rCumul.Checked then
   begin
   if GetPage = ChoixFormule  then result := nil ;
   if GetPage = ChoixSource   then result := ChoixFormule ;
   if GetPage = Budget        then result := ChoixSource ;
   if GetPage = CptVariant    then if IsBudJal then result := Budget else result := ChoixSource ;
   if GetPage = ChoixBalSit   then
     BEGIN
     If IsRubrique Then Result:=CptVariant Else result := ChoixSource ;
     END ;
   if GetPage = ChoixRec      then
     BEGIN
     If IsBalSit Then
       BEGIN
       (*
       If IsRubrique Then Result:=ChoixBalSit Else result := ChoixSource ;
       *)
       Result:=ChoixBalSit ;
       END Else
       BEGIN
       If IsRubrique Then Result:=CptVariant Else
         if IsBudJal then result := Budget else result := ChoixSource ;
       END ;
     END ;
   if GetPage = ChoixRec2     then Result:=ChoixRec ;
   if GetPage = Criteres      then
     BEGIN
     If IsBalSit Then Result:=ChoixRec Else
       BEGIN
       If RSource.ItemIndex In [10..12] Then Result:=ChoixRec2 Else Result:=ChoixRec ;
       END ;
     END ;
   if GetPage = Etablissement then
     BEGIN
     If IsBalSit Then Result:=ChoixRec Else
       BEGIN
       if not IsBudgete then result := Criteres else result := ChoixRec ;
       END ;
     END ;
   if GetPage = Devise        then
     BEGIN
     If IsBalSit Then Result:=ChoixRec Else result := Etablissement ;
     END ;
   if GetPage = Periode       then
     BEGIN
     If IsBalSit Then Result:=ChoixRec Else result := Devise ;
     END ;
   if GetPage = Pavances      then
     BEGIN
     If IsBalSit Then Result:=ChoixRec Else result := Periode ;
     END ;
   if GetPage = Resultat      then
     BEGIN
     If OkAvances(FALSE)<>avNone Then Result:=PAvances Else result := Periode ;
     END ;
   end ;
if rChamp.Checked then
   begin
   if GetPage = ChoixFormule  then result := nil ;
   if GetPage = ChoixSource   then result := ChoixFormule ;
   if GetPage = ChoixRec      then result := ChoixSource ;
   if GetPage = ChoixChamp    then if not IsSociete then result := ChoixRec else result := ChoixSource ;
   if GetPage = Resultat      then result := ChoixChamp ;
   end ;
if rConstante.Checked then
   begin
   if GetPage = ChoixFormule  then result := nil ;
   if GetPage = ChoixRec      then result := ChoixFormule ;
   if GetPage = Etablissement then result := ChoixRec ;
   if GetPage = Periode       then result := Etablissement ;
   if GetPage = Resultat      then
      begin
      if bLib.Down then result := Etablissement else result := Periode ;
      end ;
   end ;
if rSQL.Checked then
   begin
   if GetPage = ChoixFormule  then result := nil ;
   if GetPage = PSQL          then result := ChoixFormule ;
   if GetPage = Resultat      then result := PSQL ;
   end ;
end ;

function  TFAssistOLE.NextPage : TTabSheet ;
begin
result := nil ;
if rCumul.Checked then
   begin
   if GetPage = ChoixFormule  then result := ChoixSource ;
   if GetPage = ChoixSource   then
     BEGIN
     if IsBudJal then result := Budget else
       If IsRubrique Then Result:=CptVariant Else
         If IsBalSit Then Result:=ChoixBalSit Else result := ChoixRec ;
     END ;
   if GetPage = Budget        then If IsRubrique Then Result:=CptVariant Else result := ChoixRec ;
   if GetPage = CptVariant    then
     BEGIN
     If IsBalSit Then Result:=ChoixBalSit Else result := ChoixRec ;
     END ;
   if GetPage = ChoixBalSit Then
     BEGIN
     result := ChoixRec ;
     END ;
   if GetPage = ChoixRec      then
     BEGIN
     If IsBalSit Then Result:=PAvances Else
       BEGIN
       If RSource.ItemIndex In [10..12] Then Result:=ChoixRec2 Else
         BEGIN
         if not IsBudgete then result := Criteres else result := Etablissement ;
         END ;
       END ;
     END ;
   if GetPage = ChoixRec2      then
     BEGIN
     if not IsBudgete then result := Criteres else result := Etablissement ;
     END ;
   if GetPage = Criteres      then result := Etablissement ;
   if GetPage = Etablissement then result := Devise ;
   if GetPage = Devise        then result := Periode ;
   if GetPage = Periode       then
     BEGIN
     If OkAvances(FALSE)<>avNone Then Result:=PAvances Else result := Resultat ;
     END ;
   if GetPage = PAvances      then result := Resultat ;
   if GetPage = Resultat      then result := nil ;
   end ;
if rChamp.Checked then
   begin
   if GetPage = ChoixFormule  then result := ChoixSource ;
   if GetPage = ChoixSource   then if IsSociete then result := ChoixChamp else result := ChoixRec ;
   if GetPage = ChoixRec      then result := ChoixChamp ;
   if GetPage = ChoixChamp    then result := Resultat ;
   if GetPage = Resultat      then result := nil ;
   end ;
if rConstante.Checked then
   begin
   if GetPage = ChoixFormule  then result := ChoixRec ;
   if GetPage = ChoixRec      then result := Etablissement ;
   if GetPage = Etablissement then
      begin
      if bLib.Down then result := Resultat else result := Periode ;
      end ;
   if GetPage = Periode       then result := Resultat ;
   if GetPage = Resultat      then result := nil ;
   end ;
if rSQL.Checked then
   begin
   if GetPage = ChoixFormule  then result := PSQL ;
   if GetPage = PSQL          then result := Resultat ;
   if GetPage = Resultat      then result := NIL ;
   end ;
end ;

(*
function TFAssistEF.PageCount : Integer ;
begin
result := 9 ;
if rCumul.Checked then result := 9 ;
if rChamp.Checked then result := 6 ;
if rConstante.Checked then
   begin
   if bLib.Down then result := 4 else result := 5 ;
   end ;
end ;
*)
procedure TFAssistOLE.ChangeSourceItems(n : Integer) ;
var i : Integer ;
begin
rSource.Items.Clear ;
for i:=1 to n do
   begin
//   rSource.Items.Add(MsgEF.Mess[32+i]) ;
   rSource.Items.Add(MsgEF.Mess[59+i]) ;
   end ;
rSource.ItemIndex := 0 ;
rSource.Enabled := True ;
end ;

procedure TFAssistOLE.SetInitialValues ;
var sf,st,sif : String ;
    f         : Integer ;
    Cum2      : Boolean ;
begin
LoadingFormula := True ;
sf  := InitialFormula ;
sif := InitialFormula ; // sauvegarde de la formule
delete(sf,1,1) ;
f:=0 ; Cum2:=FALSE ; 
if Uppercase(Copy(sf,1,10)) = 'GET_CUMUL2' then
  begin
  delete(sf,1,10) ; f := 1 ; Cum2:=TRUE ; GC2.Checked:=TRUE ;
  end Else
  BEGIN
  if Uppercase(Copy(sf,1,9)) = 'GET_CUMUL' then
    begin
    delete(sf,1,9) ; f := 1 ;
    end ;
  END ;
if Uppercase(Copy(sf,1,8)) = 'GET_INFO'  then begin delete(sf,1,8) ; f := 2 ; end ;
if Uppercase(Copy(sf,1,13)) = 'GET_CONSTANTE' then begin delete(sf,1,13) ; f := 3 ; end ;
if Uppercase(Copy(sf,1,7)) = 'GET_SQL' then begin delete(sf,1,7) ; f := 4 ; end ;
if f=0 then begin LoadingFormula := False ; exit ; end ; // erreur dans la formule ou pas de formule
sf := Copy(sf,2,Length(sf)-2) ;
if f=1 then // cumul
   begin
   rCumul.Checked := True ;
   rChamp.Checked := False ;
   rConstante.Checked := False ;
   rSQL.Checked:=FALSE ;
   st := ReadTokenSt(sf) ;
   SetInitialTable(st) ;
   SetInitialCptVariant(LeCptVariant) ;
   st := ReadTokenSt(sf) ;
   SetInitialRec(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialEcr(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialEtab(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialDev(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialDate(st,FALSE) ;
   If Cum2 Then
     BEGIN
     st := ReadTokenSt(sf) ;
     st := ReadTokenSt(sf) ;
     SetInitialRec2(st) ;
     st := ReadTokenSt(sf) ;
     SetInitialBalSit(st) ;
     st := ReadTokenSt(sf) ;
     SetInitialWhereSup(st) ;
     END ;
   end ;
if f=2 then // info
   begin
   rCumul.Checked := False ;
   rChamp.Checked := True ;
   rConstante.Checked := False ;
   rSQL.Checked:=FALSE ;
   st := ReadTokenSt(sf) ;
   SetInitialTable(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialRec(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialField(st) ;
   st := ReadTokenSt(sf) ;
   cLibelle.Checked := (st = '"$"') ;
   end ;
if f=3 then // constante
   begin
   rCumul.Checked := False ;
   rChamp.Checked := False ;
   rConstante.Checked := True ;
   rSQL.Checked:=FALSE ;
   bVal.Down := True ;
   st := ReadTokenSt(sf) ;
   if Copy(st,2,1) = '$' then begin Delete(st,2,1) ; bLib.Down := True ; end ;
   SetInitialRec(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialEtab(st) ;
   st := ReadTokenSt(sf) ;
   SetInitialDate(st,TRUE) ;
   end ;
if f=4 then // SQL
   begin
   rCumul.Checked := False ;
   rChamp.Checked := False ;
   rConstante.Checked := FALSE ;
   rSQL.Checked:=TRUE ;
   st := ReadTokenSt(sf) ;
   SetInitialSQL(st) ;
   end ;
// forçages car mis à False par FlagExtRec dans SetInitialRec
rSource.Enabled := True ;
cAxe.Enabled := True ;
InitialFormula := sif ; // restauration de l'ancienne formule
LoadingFormula := False ;
end ;

Function DecodeCptVariant(Var St,LeCptVariant: String) : Boolean;
Var Pos1,Pos2 : Integer ;
BEGIN
Pos1:=Pos('(',St) ; Pos2:=Pos(')',St) ;
LeCptVariant:='' ; Result:=FALSE ;
If (Pos1>0) And (Pos2>0) Then
  BEGIN
  LeCptVariant:=Copy(St,Pos1+1,Pos2-Pos1-1) ;
  St:=Copy(St,1,Pos1-1) ;
  Result:=TRUE ;
  END ;
END ;

procedure TFAssistOLE.SetInitialTable(s : String) ;
var st : String ;
    l  : integer ;
begin
st := Trim(Uppercase(s)) ;
st := Copy(st,2,length(st)-2) ; // sans les "
DecodeCptVariant(St,LeCptVariant) ;
l := Length(st) ;
if Copy(st,l-1,1)=':' then
   begin
   if Copy(st,l,1)='B' then rBudget.ItemIndex:=0 else rBudget.ItemIndex:=1 ;
   rBudgetClick(nil) ;
   st := Copy(st,1,l-2) ;
   end ;
if st = 'RUBRIQUE'         then rSource.ItemIndex := 0 ;
if st = 'RUBBUD'           then rSource.ItemIndex := 1 ;
if st = 'GENERAUX'         then rSource.ItemIndex := 2 ;
if st = 'TIERS'            then rSource.ItemIndex := 3 ;
if st = 'SECTION'          then rSource.ItemIndex := 4 ;
if st = 'BUDGENE'          then rSource.ItemIndex := 5 ;
if st = 'BUDSECT'          then rSource.ItemIndex := 6 ;
if st = 'JOURNAL'          then rSource.ItemIndex := 7 ;
if st = 'BUDJAL'           then rSource.ItemIndex := 8 ;
if st = 'SOCIETE'          then rSource.ItemIndex := 9 ;
if st = 'GENERAUX/TIERS'   then rSource.ItemIndex := 10 ;
if st = 'GENERAUX/SECTION' then rSource.ItemIndex := 11 ;
if st = 'BUDGENE/BUDSECT'  then rSource.ItemIndex := 12 ;
end ;

procedure TFAssistOLE.SetInitialRec(s : String) ;
var st : String ;
    Qi : TQuery ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtRec.Checked := True ;
   RefCelRec.Text := st ;
   bFin.Enabled := True ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
if Copy(st,1,1) = '$' then Delete(st,1,1) ; // gérer la case à cocher correspondante
if rSource.ItemIndex = 4 then // section analytique (chercher l'axe)
   begin
   Qi := OpenSQL('SELECT * FROM SECTION WHERE S_SECTION = "' + st + '"',True) ;
   if not Qi.EOF then
      begin
      cAxe.ItemIndex := StrToInt(Qi.FindField('S_AXE').AsString) + 1 ;
      cAxeChange(nil) ;
      end ;
   Ferme(Qi) ;
   end ;
if rSource.ItemIndex = 6 then // section budgétaire (chercher l'axe)
   begin
   Qi := OpenSQL('SELECT * FROM BUDSECT WHERE BS_BUDSECT = "' + st + '"',True) ;
   if not Qi.EOF then
      begin
      cAxe.ItemIndex := StrToInt(Qi.FindField('BS_AXE').AsString) + 1 ;
      cAxeChange(nil) ;
      end ;
   Ferme(Qi) ;
   end ;
InitialRec := st ;
end ;

procedure TFAssistOLE.SetInitialRec2(s : String) ;
var st : String ;
    Qi : TQuery ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtRec2.Checked := True ;
   RefCelRec2.Text := st ;
   bFin.Enabled := True ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
if Copy(st,1,1) = '$' then Delete(st,1,1) ; // gérer la case à cocher correspondante
if rSource.ItemIndex = 10 then // section analytique (chercher l'axe)
   begin
   Qi := OpenSQL('SELECT * FROM SECTION WHERE S_SECTION = "' + st + '"',True) ;
   if not Qi.EOF then
      begin
      cAxe.ItemIndex := StrToInt(Qi.FindField('S_AXE').AsString) + 1 ;
      cAxeChange(nil) ;
      end ;
   Ferme(Qi) ;
   end ;
if rSource.ItemIndex = 12 then // section budgétaire (chercher l'axe)
   begin
   Qi := OpenSQL('SELECT * FROM BUDSECT WHERE BS_BUDSECT = "' + st + '"',True) ;
   if not Qi.EOF then
      begin
      cAxe.ItemIndex := StrToInt(Qi.FindField('BS_AXE').AsString) + 1 ;
      cAxeChange(nil) ;
      end ;
   Ferme(Qi) ;
   end ;
InitialRec2 := st ;
end ;

procedure TFAssistOLE.SetInitialBalSit(s : String) ;
var st : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtBalSit.Checked := True ;
   RefCelBalSit.Text := st ;
   CCBalSit.Checked:=TRUE ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
St:=Trim(St) ;
If St<>'' Then
  BEGIN
  cBalSit.Value := st ;
  CCBalSit.Checked:=TRUE ;
  END ;
end ;

Procedure RecupChamp(St : String ; Var StCh,StOp,StVal,StEtOu : String ; Var Contient : Boolean) ;
Var i,j : Integer ;
    StOp1,StOpNot : String ;
(*
(E_AFFAIRE<'4545' AND
(E_AUXILIAIRE LIKE '6596%' OR E_BANQUEPREVI LIKE '%5654%' OR
(E_BUDGET NOT  BETWEEN '4545,1212' AND '4545,1212' AND E_CONSO NOT LIKE '545%')))
*)
BEGIN
StCh:='' ; StOp:='' ; StVal:='' ; j:=0 ; StEtOu:='' ;  StOp1:='' ; StOpNot:='' ; Contient:=FALSE ;
For i:=1 To Length(St) Do
  BEGIN
  If (St[i]=' ') or (St[i]='<') or (St[i]='>') or (St[i]='=') Then
    BEGIN j:=i ; Break ; END ;
  END ;
If j=0 Then Exit ;
StCh:=Trim(Copy(St,1,j-1)) ;
If St[j]=' ' Then delete(St,1,j) Else delete(St,1,j-1) ;
St:=Trim(St) ;
i:=Pos('NOT',St) ;
If i>0 Then BEGIN St:=FindEtReplace(St,'NOT','   ',FALSE) ; StOpNot:='NOT' ; END ;
St:=Trim(St) ;
i:=Pos('AND',St) ; If i<>0 Then BEGIN StEtOu:='AND' ; Delete(St,i,3) ; END ;
i:=Pos('OR',St) ; If i<>0 Then BEGIN StEtOu:='OR' ; Delete(St,i,2) ; END ;
i:=Pos('>=',St) ; If i<>0 Then BEGIN StOp1:='>=' ; Delete(St,i,2) ; END ;
i:=Pos('<=',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='<=' ; Delete(St,i,2) ; END ;
i:=Pos('<>',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='<>' ; Delete(St,i,2) ; END ;
i:=Pos('=',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='=' ; Delete(St,i,1) ; END ;
i:=Pos('>',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='>' ; Delete(St,i,1) ; END ;
i:=Pos('<',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='<' ; Delete(St,i,1) ; END ;
i:=Pos('BETWEEN',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='BETWEEN' ; Delete(St,i,7) ; END ;
i:=Pos('LIKE',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='LIKE' ; Delete(St,i,4) ; END ;
i:=Pos('IN',St) ; If (i<>0) And (StOp1='') Then BEGIN StOp1:='IN' ; Delete(St,i,2) ; END ;
StOp:=StOpNot+' '+StOp1 ; StOp:=Trim(StOp) ;
St:=FindEtReplace(St,'(',' ',TRUE) ;
St:=FindEtReplace(St,')',' ',TRUE) ;
St:=FindEtReplace(St,'''',' ',TRUE) ;
St:=Trim(St) ;
If St[1]='%' Then Contient:=TRUE ;
St:=FindEtReplace(St,'%',' ',TRUE) ;
St:=Trim(St) ;
StVal:=St ;
END ;

Type tii = Array[0..5] Of Integer ;

Function GetIndChampWhereSup(St : String ; av : tav ; Var pref1 : String) : Integer ;
var 
    i : Integer ;
BEGIN
Result:=0 ; pref1:='' ;
For i:=1 To length(St) Do If (St[i]='_') And (i>1) Then
  BEGIN
  Case av of
    avEcr : If St[i-1]='E' Then BEGIN Pref1:='E_' ; Result:=i-1 ; Exit ; END ;
    avAna : If St[i-1]='Y' Then BEGIN Pref1:='Y_' ; Result:=i-1 ; Exit ; END ;
    avCpt : If St[i-1]='G' Then BEGIN Pref1:='G_' ; Result:=i-1 ; Exit ; END Else
             If St[i-1]='T' Then BEGIN Pref1:='T_' ; Result:=i-1 ; Exit ; END Else
              If St[i-1]='S' Then BEGIN Pref1:='S_' ; Result:=i-1 ; Exit ; END ;
    End ;
  END ;
END ;

Procedure TFAssistOLE.DecodeWhereSup(S : String) ;
Var St,St1 : String ;
    i,j,Max,Ind : integer ;
    Pref,Pref1 : String ;
    ii : tii ;
    StCh,StOp,StVal,StEtOu : String ;
    Z_C,ZO : THValComboBox ;
    ZV : tEdit ;
    ZG : tComboBox ;
    Contient : Boolean ;
    ForceIndex : Integer ;
    av : tav ;
(*
(E_AFFAIRE<'4545' AND
(E_AUXILIAIRE LIKE '6596%' OR
E_BANQUEPREVI LIKE '%5654%' OR
(E_BUDGET NOT  BETWEEN '4545,1212'
AND '4545,1212' AND
E_CONSO NOT LIKE '545%')))
*)
BEGIN
St:=Trim(S) ; If St='' Then Exit ;
For i:=length(St) DownTo 1 Do
  BEGIN
  If St[i]=')' Then St[i]:=' ' Else Break ;
  END ;
St:=Trim(St) ;
Fillchar(ii,SizeOf(ii),#0) ;
ForceIndex:=-1 ;
If (Pos('E_',St)>0) Then ForceIndex:=0 ;
If (Pos('Y_',St)>0) Then ForceIndex:=1 ;
If ForceIndex=-1 Then
  BEGIN
  If (Pos('G_',St)>0) Then ForceIndex:=0 ;
  If (Pos('T_',St)>0) Then ForceIndex:=0 ;
  If (Pos('S_',St)>0) Then ForceIndex:=1 ;
  END ;
//Case OkAvances(TRUE) Of avEcr : Pref:='E_' ; avAna : pref:='Y_' else exit ; End ;
av:=OkAvances(TRUE,ForceIndex) ;
Case av Of avEcr : Pref:='E_' ; avAna : pref:='Y_'  ; avCpt : Pref:='E' ; else exit ; End ;
//ChargeAvances(FALSE) ;
(*
St1:=St ; j:=0 ;
While Pos(Pref,St1)>0 Do
  BEGIN
  ii[j]:=Pos(Pref,St1) ; inc(j) ;
  St1:=FindEtReplace(St1,Pref,'!!',FALSE) ;
  END ;
Max:=j-1 ;
*)
St1:=St ; j:=0 ;
While GetIndChampWhereSup(St1,av,Pref1)>0 Do
  BEGIN
  ii[j]:=GetIndChampWhereSup(St1,av,pref1) ; inc(j) ;
  St1:=FindEtReplace(St1,Pref1,'!!',FALSE) ;
  END ;
Max:=j-1 ;
For Ind:=0 To Max Do
  BEGIN
  if Ind<>Max Then St1:=Copy(St,ii[Ind],ii[ind+1]-ii[Ind])
              Else St1:=Copy(St,ii[Ind],Length(St)-ii[Ind]+1);
  St1:=FindEtReplace(St1,'(',' ',TRUE) ;
  St1:=FindEtReplace(St1,')',' ',TRUE) ;
  St1:=Trim(St1) ;
  RecupChamp(St1,StCh,StOp,StVal,StEtOu,Contient) ;
  If StCh<>'' Then
    BEGIN
    Z_C:=THvalComboBox(FindComponent('Z_C'+IntToStr(Ind+1))) ;
    ZO:=THvalComboBox(FindComponent('ZO'+IntToStr(Ind+1))) ;
    ZV:=TEdit(FindComponent('ZV'+IntToStr(Ind+1))) ;
    ZG:=THvalComboBox(FindComponent('ZG'+IntToStr(Ind+1))) ;
    If Z_C<>NIL Then Z_C.Value:=StCh ;
    If ZO<>NIL Then
      BEGIN
      If StOp='<' Then ZO.Value:='<' ;
      If StOp='<=' Then ZO.Value:='<=' ;
      If StOp='<>' Then ZO.Value:='<>' ;
      If StOp='=' Then ZO.Value:='=' ;
      If StOp='>' Then ZO.Value:='>' ;
      If StOp='>=' Then ZO.Value:='>=' ;
      If Contient Then
         BEGIN
         If StOp='LIKE' Then ZO.Value:='L' ;
         If StOp='NOT LIKE' Then ZO.Value:='M' ;
         END Else
         BEGIN
         If StOp='LIKE' Then ZO.Value:='C' ;
         If StOp='NOT LIKE' Then ZO.Value:='D' ;
         END ;
      If StOp='BETWEEN' Then ZO.Value:='E' ;
      If StOp='NOT BETWEEN' Then ZO.Value:='G' ;
      If StOp='IN' Then ZO.Value:='I' ;
      If StOp='NOT IN' Then ZO.Value:='J' ;
      END ;
    If ZV<>NIL Then ZV.Text:=StVal ;
    If ZG<>NIL Then
      BEGIN
      If StEtOu='AND' Then ZG.Itemindex:=0 Else If StEtOu='OR' Then ZG.Itemindex:=1 ;
      END ;
    END ;
  END ;

END ;

procedure TFAssistOLE.SetInitialWhereSup(s : String) ;
Var St : String ;
BEGIN
st := Trim(Uppercase(s)) ;
(*
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtEtab.Checked := True ;
   RefCelEtab.Text := st ;
   exit ;
   end ;
*)
st := Copy(st,2,length(st)-2) ; // sans les "
If St='' Then GC2.Checked:=TRUE Else DecodeWhereSup(St) ;
END ;

procedure TFAssistOLE.SetInitialField(s : String) ;
var st : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtField.Checked := True ;
   RefCelField.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
InitialField := st ;
end ;

procedure TFAssistOLE.SetInitialEcr(s : String) ;
var st,sr : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtEcr.Checked := True ;
   RefCelEcr.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
if Copy(st,Length(st),1) = 'R' then
   begin
   Delete(st, Length(st), 1) ;
   Revision.ItemIndex := 0 ;
   end
else Revision.ItemIndex := 1 ;
if Pos('N',st)>0 then
   begin
   if Pos('NSU',st)>0 then sr := 'NSS' else sr := 'NOR' ;
   if Pos('NSUP',st)>0 then sr := '' ;
   end
else
   begin
   if Pos('SU',st)>0 then sr := 'SSI' ;
   if Pos('P',st)>0 then sr := 'PRE' ;
   end ;
cNatureEcr.Value := sr ;
cIntegAN.ItemIndex := 0 ;
AvecIFRS.Checked := False ;
if Pos('-',st)>0 then cIntegAN.ItemIndex := 1 ;
if Pos('#',st)>0 then cIntegAN.ItemIndex := 2 ;
if cNatureEcr.ItemIndex = -1 then cNatureEcr.ItemIndex := 0 ;
if cIntegAN.ItemIndex = -1 then cIntegAN.ItemIndex := 0 ;
CSens.Value:='SM' ;
If Pos('D',St)>0 Then CSens.Value:='TD' Else
 If Pos('C',St)>0 Then CSens.Value:='TC' Else
  If Pos('/',St)>0 Then CSens.Value:='SD' Else
   If Pos('\',St)>0 Then CSens.Value:='SC' ;
CMonnaie.ItemIndex:=0 ;
If Pos('O',St)>0 Then CMonnaie.ItemIndex:=1 ;
end ;

procedure TFAssistOLE.SetInitialEtab(s : String) ;
var st : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtEtab.Checked := True ;
   RefCelEtab.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
cEtablissement.Value := st ;
if cEtablissement.ItemIndex = -1 then cEtablissement.ItemIndex := 0 ;
end ;

procedure TFAssistOLE.SetInitialSQL(s : String) ;
var st : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtEtab.Checked := True ;
   RefCelEtab.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
InitialSQL:=St ;
Z_SQL.lines.add(InitialSQL);
//cEtablissement.Value := st ;
end ;

procedure TFAssistOLE.SetInitialDev(s : String) ;
var st : String ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtDev.Checked := True ;
   RefCelDev.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
cDevise.Value := st ;
if cDevise.ItemIndex = -1 then cDevise.ItemIndex := 0 ;
end ;

procedure TFAssistOLE.SetInitialCptVariant(s : String) ;
var st : String ;
    i : Integer ;
begin
st := Trim(Uppercase(s)) ;
if Pos('"',st) > 0 then // référence cellule
   begin
   FlagExtVariant.Checked := True ;
   While Pos('&',St)>0 Do BEGIN i:=Pos('&',St) ; Delete(St,i,1) ; END ;
   While Pos('"',St)>0 Do BEGIN i:=Pos('"',St) ; Delete(St,i,1) ; END ;
   RefCellVariant.Text := st ;
   exit ;
   end ;
//st := Copy(st,2,length(st)-2) ;
cCptVariant.Text := st ;
end ;


procedure TFAssistOLE.SetInitialDate(s : String ; PourConstante : Boolean) ;
var st,StNum1,StNum2,StNum,StY : String ;
    i : Integer ;
begin
st := Trim(Uppercase(s)) ; If Pos('(',St)>0 Then Exit ;
if Pos('"',st) = 0 then // référence cellule
   begin
   FlagExtPer.Checked := True ;
   RefCelPer.Text := st ;
   exit ;
   end ;
st := Copy(st,2,length(st)-2) ; // sans les "
cPeriode.Value := Copy(st,1,1) ; // N,B,T,Q,S ...
cExercice.ItemIndex := 0 ;
If (CPeriode.Value='W') And (Not PourConstante) Then
  BEGIN
  delete(st,1,2) ; StNum:='1' ; StY:='2000' ;
  I:=Pos(':',St) ; If i>0 Then BEGIN StNum:=Copy(St,1,i-1) ; StY:=Copy(St,i+1,Length(St)-I) ; END ;
  i:=Pos('-',StNum) ; If i>0 Then BEGIN StNum1:=Copy(StNum,1,i-1) ; StNum2:=Copy(StNum,i+1,Length(StNum)-I) ; END ;
  NumPeriode.Value:=StrToInt(StNum1) ; NumPeriode1.Value:=StrToInt(StNum2) ;
  If Length(St)<=2 Then BEGIN If StrToInt(StY)>95 Then StY:='19'+StY Else StY:='20'+StY ; END ;
  cAnnee.Value:=StY ;
  END Else
  BEGIN
  delete(st,1,1) ;
  if pos('---',st) > 0 then begin cExercice.ItemIndex := 3 ; delete(st,Pos('---',st),3) ; end ;
  if pos('--',st) > 0 then begin cExercice.ItemIndex := 2 ; delete(st,Pos('--',st),2) ; end ;
  if pos('-',st) > 0 then begin cExercice.ItemIndex := 1 ; delete(st,Pos('-',st),1) ; end ;
  if pos('+',st) > 0 then begin cExercice.ItemIndex := 4 ; delete(st,Pos('+',st),1) ; end ;
  cDetailPeriode.ItemIndex := 0 ;
  if pos('<',st) > 0 then begin cDetailPeriode.ItemIndex := 1 ; delete(st,Pos('<',st),1) ; end ;
  if pos('>',st) > 0 then begin cDetailPeriode.ItemIndex := 2 ; delete(st,Pos('>',st),1) ; end ;
  if st <> '' then
     begin
     cNumPeriode.ItemIndex := StrToInt(st)-1 ;
     NumPeriode.Value := StrToInt(st) ;
     end ;
  if cPeriode.ItemIndex = -1 then cPeriode.ItemIndex := 0 ;
  if cNumPeriode.ItemIndex = -1 then cNumPeriode.ItemIndex := 0 ;
  if cExercice.ItemIndex = -1 then cExercice.ItemIndex := 0 ;
  END ;
end ;

{ Evènements de la form }

procedure TFAssistOLE.ChangementCode(var Grille : THGrid; FindCod, FindLibelle : string);
Var
SQL, Latable   : string;
QLe            : TQuery;
TobOle         : TOB;
begin
inherited;
if rConstante.Checked then
   begin
   SQL := 'SELECT CC_CODE Code, CC_LIBELLE Libelle FROM CHOIXCOD WHERE CC_TYPE = "CON" AND CC_CODE LIKE "' + FindCod + '%" ORDER BY CC_CODE'; Latable := 'CHOIXCOD';
   end
else
   begin
   Case rSource.ItemIndex of
      0 : begin
              SQL := 'SELECT RB_RUBRIQUE Code , RB_LIBELLE Libelle FROM RUBRIQUE WHERE RB_NATRUB="CPT" ';
              if FindCod <> '' then SQL := SQL + ' AND RB_RUBRIQUE LIKE "' + FindCod + '%"' ;
              if FindLibelle <> '' then SQL := SQL +  ' AND RB_LIBELLE LIKE "%' + FindLibelle + '%"';
              SQL := SQL + ' ORDER BY RB_RUBRIQUE'; Latable := 'RUBRIQUE';
          end;
      1 : begin
              SQL := 'SELECT RB_RUBRIQUE Code, RB_LIBELLE Libelle FROM RUBRIQUE WHERE ' + WhereBudJal + 'RB_NATRUB="BUD" ';
              if FindCod <> '' then SQL := SQL + ' AND RB_RUBRIQUE LIKE "' + FindCod + '%"';
              if FindLibelle <> '' then SQL := SQL + ' AND RB_LIBELLE LIKE "%' + FindLibelle + '%"';
              SQL := SQL + ' ORDER BY RB_RUBRIQUE';
              Latable := 'RUBRIQUE';
          end;
      2,10,11,12 : begin
              SQL := 'SELECT G_GENERAL Code, G_LIBELLE Libelle FROM GENERAUX ';
              if (FindCod <> '') and (FindLibelle <> '') then
                SQL := SQL + ' WHERE G_GENERAL LIKE "' + FindCod + '%" AND G_LIBELLE LIKE "' + FindLibelle + '%"'
              else if (FindCod <> '') then
                SQL := SQL + ' WHERE G_GENERAL LIKE "' + FindCod + '%"'
              else if FindLibelle <> '' then
                SQL := SQL +  'WHERE G_LIBELLE LIKE "' + FindLibelle + '%"';
              SQL := SQL + ' ORDER BY G_GENERAL'; Latable := 'GENERAUX';
          end;
      3 : begin
              SQL := 'SELECT T_AUXILIAIRE Code, T_LIBELLE Libelle FROM TIERS ';
              if (FindCod <> '') and (FindLibelle <> '') then
              SQL := SQL +  ' WHERE T_AUXILIAIRE LIKE "' + FindCod + '%"' + ' AND  T_LIBELLE LIKE "' + FindLibelle + '%"'
              else
              if (FindCod <> '') then
              SQL := SQL +  ' WHERE T_AUXILIAIRE LIKE "' + FindCod + '%"'
              else
              if (FindLibelle <> '') then
              SQL := SQL +  ' WHERE  T_LIBELLE LIKE "' + FindLibelle + '%"';
              SQL := SQL +  ' ORDER BY T_AUXILIAIRE'; Latable := 'TIERS';
      end;
      4 : begin
              SQL := 'SELECT S_SECTION Code , S_LIBELLE Libelle FROM SECTION WHERE S_AXE = "' + AxeCourant + '"';
              if (FindCod <> '') then SQL := SQL + ' AND S_SECTION LIKE "' + FindCod + '%"';
              if (FindLibelle <> '') then SQL := SQL +  ' AND  S_LIBELLE LIKE "'+ FindLibelle + '%"';
              SQL := SQL + ' ORDER BY S_SECTION';  Latable := 'SECTION';
          end;
      5 : begin
              SQL := 'SELECT BG_BUDGENE Code, BG_LIBELLE Libelle FROM BUDGENE ';
              if (FindCod <> '') and (FindLibelle <> '') then
              SQL := SQL + ' WHERE BG_BUDGENE LIKE "' + FindCode.Text + '%" ' + ' AND BG_LIBELLE LIKE "' + FindLib.Text + '%" '
              else
              if (FindCod <> '') then SQL := SQL + ' WHERE BG_BUDGENE LIKE "' + FindCod + '%" '
              else
              if (FindLibelle <> '') then SQL := SQL + ' WHERE BG_LIBELLE LIKE "' + FindLibelle + '%" ';
              SQL := SQL + ' ORDER BY BG_BUDGENE'; Latable := 'BUDGENE';
          end;
      6 : begin
              SQL := 'SELECT BS_BUDSECT Code, BS_LIBELLE Libelle FROM BUDSECT WHERE BS_AXE = "' + AxeCourant+ '%"';
              if (FindCod <> '') then SQL := SQL + '" AND BS_BUDSECT LIKE "' + FindCod + '%"';
              if (FindLibelle <> '') then SQL := SQL + '" AND LIBELLE LIKE "' + FindLibelle + '%"';
              SQL := SQL + ' ORDER BY BS_BUDSECT';  Latable := 'BUDSECT';
          end;
      7 : begin
              SQL := 'SELECT J_JOURNAL Code, J_LIBELLE Libelle FROM JOURNAL ';
              if (FindCod <> '') and (FindLibelle <> '') then
              SQL := SQL + ' WHERE J_JOURNAL LIKE "' + FindCod + '%"'+ ' AND J_LIBELLE LIKE "' + Findlib.Text + '%"'
              else
              if (FindCod<> '') then
              SQL := SQL + ' WHERE J_JOURNAL LIKE "' + FindCod + '%"'
              else
              if (FindLibelle <> '') then
              SQL := SQL + ' WHERE J_LIBELLE LIKE "' + Findlibelle + '%"';
              SQL := SQL + ' ORDER BY J_JOURNAL'; Latable := 'JOURNAL';
          end;
      8 : begin
              SQL := 'SELECT BJ_BUDJAL Code, BJ_LIBELLE Libelle FROM BUDJAL ';
              if (FindCod <> '') and (FindLibelle <> '') then
              SQL := SQL + ' WHERE BJ_BUDJAL LIKE "' + FindCod + '%"'+ ' AND BJ_LIBELLE LIKE "' + FindLibelle + '%"'
              else
              if (FindCod <> '') then
              SQL := SQL + ' WHERE BJ_BUDJAL LIKE "' + FindCod + '%"'
              else
              if (FindLibelle <> '') then
              SQL := SQL + ' WHERE BJ_LIBELLE LIKE "' + FindLibelle + '%"';
              SQL := SQL + ' ORDER BY BJ_BUDJAL'; Latable := 'BUDJAL';
          end;
      end ;
   end ;
TobOle := TOB.Create('', nil, -1);
QLe := OpenSQL(SQL,TRUE);
TobOle.LoadDetailDB('Latable', '', '', QLe, TRUE, FALSE);
Ferme(QLe);
Grille.VidePile(False);
TobOle.PutGridDetail(Grille,False,False,'Code;Libelle');
TobOle.free;

end;

procedure TFAssistOLE.cPeriodeChange(Sender: TObject);
var i,np : Integer ;
    AA,MM,DD : Word ;
begin
inherited;
np := cNumPeriode.ItemIndex ;
cDetailPeriode.Visible := (cPeriode.ItemIndex = 1) ;
lDetailPeriode.Visible := cDetailPeriode.Visible ;
if not cDetailPeriode.Visible then cDetailPeriode.ItemIndex := 0 ;
CheckExoOKClick(nil) ;
cNumPeriode.Items.Clear ;
if cPeriode.ItemIndex <= 0 then
   begin
   NumPeriode.Text := '0' ;
   cDetailPeriode.ItemIndex := 0 ;
   exit ;
   end ;
Case cPeriode.ItemIndex of
   1 : for i:=1 to 12 do cNumPeriode.Items.Add(GetMonth(i)) ;
   2 : begin
       cNumPeriode.Items.Add(GetTwoMonth(1,2)) ;
       cNumPeriode.Items.Add(GetTwoMonth(3,4)) ;
       cNumPeriode.Items.Add(GetTwoMonth(5,6)) ;
       cNumPeriode.Items.Add(GetTwoMonth(7,8)) ;
       cNumPeriode.Items.Add(GetTwoMonth(9,10)) ;
       cNumPeriode.Items.Add(GetTwoMonth(11,12)) ;
       end ;
   3 : begin
       cNumPeriode.Items.Add(GetTwoMonth(1,3)) ;
       cNumPeriode.Items.Add(GetTwoMonth(4,6)) ;
       cNumPeriode.Items.Add(GetTwoMonth(7,9)) ;
       cNumPeriode.Items.Add(GetTwoMonth(10,12)) ;
       end ;
   4 : begin
       cNumPeriode.Items.Add(GetTwoMonth(1,4)) ;
       cNumPeriode.Items.Add(GetTwoMonth(5,8)) ;
       cNumPeriode.Items.Add(GetTwoMonth(9,12)) ;
       end ;
   5 : begin
       cNumPeriode.Items.Add(GetTwoMonth(1,6)) ;
       cNumPeriode.Items.Add(GetTwoMonth(7,12)) ;
       end ;
   6 : BEGIN //Semaine
       If NumPeriode.Value=0 Then NumPeriode.Value:=1 ;
       If NumPeriode1.Value=0 Then NumPeriode1.Value:=NumPeriode.Value ;
       DecodeDate(VH^.Entree.Deb,AA,MM,DD) ; CAnnee.Value:=FormatFloat('0000',AA) ;
       END ;
   end ;
If CPeriode.ItemIndex=6 Then
  BEGIN
  If (NumPeriode.Value<0) Or (NumPeriode.Value>53) Then NumPeriode.Value := 1 ;
  NumPeriode.Max := 53 ; NumPeriode1.Max := 53 ;
  END Else
  BEGIN
  if (cNumPeriode.ItemIndex = 1) then NumPeriode.Max := 12 ;
  if (cNumPeriode.ItemIndex = 2) then NumPeriode.Max := 6 ;
  if (cNumPeriode.ItemIndex = 3) then NumPeriode.Max := 4 ;
  if (cNumPeriode.ItemIndex = 4) then NumPeriode.Max := 3 ;
  if (cNumPeriode.ItemIndex = 5) then NumPeriode.Max := 2 ;
  if NumPeriode.Value > NumPeriode.Max then NumPeriode.Text := '1' ;
  END ;
if cNumPeriode.Visible then
   begin
   if cNumPeriode.ItemIndex = -1 then cNumPeriode.ItemIndex := np ;
   if cNumPeriode.ItemIndex = -1 then cNumPeriode.ItemIndex := 0 ;
   cNumPeriodeClick(nil) ;
   end ;
end;

procedure TFAssistOLE.cNatureEcrEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[1] ;
end;

procedure TFAssistOLE.cIntegANEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[2] ;
end;

procedure TFAssistOLE.cExerciceEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[3] ;
end;

procedure TFAssistOLE.cPeriodeEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[4] ;
end;

procedure TFAssistOLE.NumPeriodeEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[5] ;
end;

procedure TFAssistOLE.cDetailPeriodeEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[6] ;
end;

procedure TFAssistOLE.cEtablissementEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[7] ;
end;

procedure TFAssistOLE.cDeviseEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[8] ;
end;

procedure TFAssistOLE.GridRubEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[9] ;
end;

procedure TFAssistOLE.FindCodeEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[10] ;
end;

procedure TFAssistOLE.FindLibEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[11] ;
end;

procedure TFAssistOLE.bFinClick(Sender: TObject);
begin
inherited;
If WriteFormule Then ModalResult:=mrOk ;
end;

Function TFAssistOLE.WriteFormule : Boolean  ;
var s,sm,sm1,SMonnaie : String ;
    Cum2 : Boolean;
    MS : String ;
    stMS : String ;
begin
//Initialisation variables
MS := '';
stMS := '';

Result:=TRUE ;
SMonnaie:='' ;

if (RCumul.Checked) then
begin
    if (IsBalSit) Then GC2.Checked := true;
end;
Cum2 := (GC2.Checked And GC2.Visible);

if rCumul.Checked then
   begin
   //si Multidossier on rempli les variables automatique
   if MULTIDOSSIER.Value<>'' then
   begin
    MS := 'MS' ;
    stMS := ';"'+ MULTIDOSSIER.Value + '"';
   end;
   if (Cum2) then sFormule := '=Get_Cumul2' + MS + '("'
     else sFormule := '=Get_Cumul' + MS + '("';


   Case rSource.ItemIndex of
      0 : sFormule := sFormule + 'RUBRIQUE' ;
      1 : sFormule := sFormule + 'RUBBUD' ;
      2 : sFormule := sFormule + 'GENERAUX' ;
      3 : sFormule := sFormule + 'TIERS' ;
      4 : sFormule := sFormule + 'SECTION' ;
      5 : sFormule := sFormule + 'BUDGENE' ;
      6 : sFormule := sFormule + 'BUDSECT' ;
      7 : sFormule := sFormule + 'JOURNAL' ;
      8 : sFormule := sFormule + 'BUDJAL' ;
      10 : sFormule := sFormule + 'GENERAUX/TIERS' ;
      11 : sFormule := sFormule + 'GENERAUX/SECTION' ;
      12 : sFormule := sFormule + 'BUDGENE/BUDSECT' ;
      end ;
   if IsBudgetaire then
      if rBudget.ItemIndex=0 then sFormule := sFormule + ':B' else sFormule := sFormule + ':R' ;
   If IsRubrique Then
     BEGIN
     If FlagExtVariant.Checked Then SFormule:=SFormule+'("&'+RefCellVariant.Text+'&")' Else
       If Trim(cCptVariant.Text)<>'' then SFormule:=sformule+'('+cCptVariant.Text+')' ;
     END ;
   sFormule := sFormule + '"; ' ;
   if IsSection then sm := AxeCourant else sm := '' ;
   if FlagExtRec.Checked then s := RefCelRec.Text else
     BEGIN
     If FindCode.Text<>'' Then s := '"' + sm + FindCode.Text + '"'
                          Else s := '"' + sm + THGRID.Cells[0, THGRID.Row] + '"' ;
     END ;
   sFormule := sFormule + s + '; ' ;

   If CMonnaie.ItemIndex=1 Then  SMonnaie:='O' ;
   If IsCptGeneral Then
     BEGIN
     if FlagExtEcr.Checked then s := RefCelEcr.Text else
       BEGIN
       If CSens.Value<>'SM' Then
         BEGIN
         If CSens.Value='SD' Then s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ '/"' Else
          If CSens.Value='SC' Then s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ '\"' Else
           If CSens.Value='TD' Then s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ 'D"' Else
            If CSens.Value='TC' Then s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ 'C"' Else
         END Else s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ '"' ;
       END ;
     END Else
     BEGIN
     if FlagExtEcr.Checked then s := RefCelEcr.Text else s := '"' + GetNatureEcr(cNatureEcr.Value) + cIntegAN.Value +SMonnaie+ '"' ;
     END ;
   sFormule := sFormule + s + '; ' ;
   if FlagExtEtab.Checked then s := RefCelEtab.Text else s := '"' + cEtablissement.Value + '"' ;
   sFormule := sFormule + s + '; ' ;
   if FlagExtDev.Checked then s := RefCelDev.Text else s := '"' + cDevise.Value + '"' ;
   sFormule := sFormule + s + '; ' ;
   sm := Trim(NumPeriode.Text) ; if (sm = '0') or (cPeriode.ItemIndex = 0) then sm := '' ;
   if FlagExtPer.Checked then s := RefCelPer.Text else
     BEGIN
     If cPeriode.Values[cPeriode.ItemIndex]='W' Then
       BEGIN
       sm := Trim(NumPeriode.Text) ; if (sm = '0') then sm := '1' ;
       sm1 := Trim(NumPeriode1.Text) ; if (sm1 = '0') then sm1 := '1' ;
       s := '"' + cPeriode.Values[cPeriode.ItemIndex]+':'+sm+'-'+sm1+':'+Cannee.Value+ '"' ;
       END Else
       BEGIN
       s := '"' + cPeriode.Values[cPeriode.ItemIndex] + sm + cExercice.Value + cDetailPeriode.Value + '"' ;
       END ;
     END ;
   sFormule := sFormule + s + '; ' ;
   if IsBudgetaire then
      if FlagExtBud.Checked then s := RefCelBud.Text else s := '"' + cBudget.Values[cBudget.ItemIndex] + '"'
      else s := '""' ;
   If Cum2 Then
     BEGIN
     sFormule := sFormule + s + '; ' ;
     If rSource.ItemIndex in [10..12] Then
       BEGIN
       if IsSection2 then sm := AxeCourant else sm := '' ;
       if FlagExtRec2.Checked then s := RefCelRec2.Text else s := '"' + sm + GridRub2.Cells[0, THGRID.Row] + '"' ;
       END Else s := '""' ;
     If IsBalSit Then
       BEGIN
       if FlagExtBalSit.Checked then s := s+RefCelBalSit.Text else s := s+'; "' + cBalSit.Value + '"' ;
       END Else s:=s+'; ""' ;
     If OkAvances(FALSE)<>avNone Then
       BEGIN
       s:=s+'; "'+FaitSQLSupp+'"' ;
       END Else s:=s+'; ""' ;
     sFormule := sFormule + s + stMS + ')' ;
       END Else
     BEGIN
     sFormule := sFormule + s + stMS + ')' ;
     END ;
   end ;

if rChamp.Checked then
   begin
   sFormule := '=Get_Info("' ;
   Case rSource.ItemIndex of
      0 : sFormule := sFormule + 'RUBRIQUE"; ' ;
      1 : sFormule := sFormule + 'RUBBUD"; ' ;
      2 : sFormule := sFormule + 'GENERAUX"; ' ;
      3 : sFormule := sFormule + 'TIERS"; ' ;
      4 : sFormule := sFormule + 'SECTION"; ' ;
      5 : sFormule := sFormule + 'BUDGENE"; ' ;
      6 : sFormule := sFormule + 'BUDSECT"; ' ;
      7 : sFormule := sFormule + 'JOURNAL"; ' ;
      8 : sFormule := sFormule + 'BUDJAL"; ' ;
      9 : sFormule := sFormule + 'SOCIETE"; ' ;
      end ;
   if not IsSociete then
      begin
      if IsSection then sm := AxeCourant else sm := '' ;
      if FlagExtRec.Checked then s := RefCelRec.Text else s := '"' + sm + THGRID.Cells[0, THGRID.Row]  + '"' ;
      sFormule := sFormule + s + '; ' ;
      end
   else // société
      begin
      sFormule := sFormule + '"'+ V_PGI.CodeSociete + '"; ' ;
      end ;
   if FlagExtField.Checked then s := RefCelField.Text else s := '"' + GridFields.Cells[0, GridFields.Row] + '"' ;
   sFormule := sFormule + s + '; ' ;
   if cLibelle.Checked then sFormule := sFormule + '"$")' else sFormule := sFormule + '"")' ;
   end ;

if rConstante.Checked then
   begin
   sFormule := '=Get_Constante(' ;
   if bLib.Down then sm := '$' + THGRID.Cells[0, THGRID.Row] else sm := THGRID.Cells[0, THGRID.Row] ;
   if FlagExtRec.Checked then s := RefCelRec.Text else s := '"' + sm + '"' ;
   sFormule := sFormule + s + '; ' ;
   if FlagExtEtab.Checked then s := RefCelEtab.Text else s := '"' + cEtablissement.Value + '"' ;
   sFormule := sFormule + s + '; ' ;
   sm := Trim(NumPeriode.Text) ; if (sm = '0') or (cPeriode.ItemIndex = 0) then sm := '' ;
   if FlagExtPer.Checked then s := RefCelPer.Text else s := '"' + cPeriode.Value + sm + cExercice.Value + cDetailPeriode.Value + '"' ;
   sFormule := sFormule + s + ')' ;
   end ;

if rSQL.Checked then
   begin
   If Trim(uppercase(Copy(Z_SQL.text,1,6)))<>'SELECT' Then
     BEGIN
     Msg.Execute(2,'','') ; sFormule:='' ; Result:=FALSE ; exit ;
     END ;
   sFormule := '=Get_SQL(' ;
   sm:=Z_SQL.Lines.text; ;
   if FlagExtSQL.Checked then s := RefCelSQL.Text else s := '"' + sm + '"' ;
   sFormule := sFormule + s + ')' ;
   end ;

end ;

procedure TFAssistOLE.ExpliqueFormule ;
begin
r7.Caption:='' ;
if rCumul.Checked then r1.Caption := MsgF.Mess[0] ; // cumul
if rChamp.Checked then r1.Caption := MsgF.Mess[1] ; // champ
if rConstante.Checked then
   begin
   if bVal.Down then r1.Caption := MsgF.Mess[2] ; // constante (valeur)
   if bLib.Down then r1.Caption := MsgF.Mess[3] ; // constante (libellé)
   end ;
// source (table)
if rConstante.Checked then r2.Caption := MsgF.Mess[17]
                      else r2.Caption := MsgF.Mess[rSource.ItemIndex + 4] ;
// source (champ)
If Not rSQL.Checked Then
  BEGIN
  if FlagExtRec.Checked then
     begin
     r2.Caption := r2.Caption + ' ' + MsgF.Mess[20] + ' ' + RefCelRec.Text ;
     end ;
  if FlagExtRec2.Checked then
     begin
     r2.Caption := r2.Caption + ' / ' + RefCelRec2.Text ;
     end
  else
     begin
     // si <> société afficher le code de l'enregistrement sélectionné
     if not IsSociete then r2.Caption := r2.Caption + ' ' + THGRID.Cells[0, THGRID.Row] ;
     // axe éventuel
     if IsSection then r2.Caption := r2.Caption + ' (' + MsgF.Mess[14] + ' ' + AxeCourant + ')' ;
     end ;
    END ;
// type d'écritures ou champ à extraire (GetInfo)
if FlagExtEcr.Checked then
   begin
   r3.Caption := MsgF.Mess[15] + ' ' + RefCelEcr.Text ;
   end
else
   begin
   r3.Caption := cNatureEcr.Items[cNatureEcr.ItemIndex] ;
   r3.Caption := r3.Caption + chr(10) + cIntegAN.Items[cIntegAN.ItemIndex] ;
   if Revision.ItemIndex=0 then r3.Caption := r3.Caption + chr(10) + MsgF.Mess[16] ;
   end ;
if rChamp.Checked then
   begin
   l3.Caption := MsgF.Mess[18] ; // champ
   if FlagExtField.Checked then r3.Caption := MsgF.Mess[15] + ' ' + RefCelField.Text
                           else r3.Caption := GridFields.Cells[0, GridFields.Row] ;
   end
else l3.Caption := MsgF.Mess[19] ; // type d'écritures
// établissement
if FlagExtEtab.Checked then
   r4.Caption := MsgF.Mess[15] + ' ' + RefCelEtab.Text
else
   r4.Caption := cEtablissement.Items[cEtablissement.ItemIndex] ;
// devise
if FlagExtDev.Checked then
   r5.Caption := MsgF.Mess[15] + ' ' + RefCelDev.Text
else
   r5.Caption := cDevise.Items[cDevise.ItemIndex] ;
// période
if FlagExtPer.Checked then
   begin
   r6.Caption := MsgF.Mess[15] + ' ' + RefCelPer.Text ;
   end
else
   begin
   r6.Caption := cExercice.Items[cExercice.ItemIndex] ;
   r6.Caption := r6.Caption + chr(10) + cPeriode.Items[cPeriode.ItemIndex] ;
   if cPeriode.ItemIndex > 0 then
      begin
      if cNumPeriode.Visible then r6.Caption := r6.Caption + ' ' + cNumPeriode.Items[cNumPeriode.ItemIndex]
                             else r6.Caption := r6.caption + ' ' + NumPeriode.Text ;
      r6.Caption := r6.Caption + chr(10) + cDetailPeriode.Items[cDetailPeriode.ItemIndex] ;
      end ;
   end ;
If rCumul.Checked And GC2.Checked Then
  BEGIN
  r7.Caption:=FaitSQLSupp ;
  END ;
r1.Visible := True ; l1.Visible := True ;
r2.Visible := True ; l2.Visible := True ;
r3.Visible := not rConstante.Checked ; l3.Visible := r3.Visible ;
r4.Visible := not rChamp.Checked ; l4.Visible := r4.Visible ;
r5.Visible := not rChamp.Checked ; l5.Visible := r5.Visible ;
r6.Visible := rCumul.Checked ; l6.Visible := r6.Visible ;
r7.Visible := rCumul.Checked And GC2.Checked ;
end ;

procedure TFAssistOLE.FormShow(Sender: TObject);
begin
inherited;
MakeZoomOLE(Handle) ;
Revision.Visible:=(V_PGI.LaSerie=S7) ;
//RRO 28012003
RSQL.visible:=not(V_PGI.Laserie=S3);
AvecIFRS.Visible := EstComptaIFRS ;
// BPY le 18/12/2003 => Fiche 13086 active cette coche quelque soit la serie !
//GC2.Visible:=not(V_PGI.Laserie=S3);
GC2.Visible := true;
// fin BPY

SauveAxe := 0 ;
(*
ChangeSourceItems(7) ; // Get_Cumul par défaut
*)
ChangeSourceItems(13) ;
Revision.Visible := V_PGI.Controleur ;
if InitialFormula <> '' then SetInitialValues Else
  BEGIN
  END ;
SourisNormale ;
end;

procedure TFAssistOLE.CheckExoOKClick(Sender: TObject);
begin
inherited;
CExercice.Visible:=TRUE ; NumPeriode1.Visible:=FALSE ; LNumPeriode1.Visible:=FALSE ; cAnnee.Visible:=FALSE ;
if cPeriode.ItemIndex = 0 then
  begin
  cNumPeriode.Visible := False ;  NumPeriode.Visible  := False ;
  end Else if cPeriode.ItemIndex = 6 then
  BEGIN
  cNumPeriode.Visible := FALSE ;  NumPeriode.Visible  := TRUE ; NumPeriode1.Visible:=TRUE ; LNumPeriode1.Visible:=TRUE ;
  CExercice.Visible:=FALSE ; cAnnee.Visible:=TRUE ;
  END else
  begin
  cNumPeriode.Visible := (CheckExoOK.Checked) and (cPeriode.ItemIndex>0) ;
  NumPeriode.Visible  := not cNumPeriode.Visible ;
  end ;
end;

procedure TFAssistOLE.cNumPeriodeClick(Sender: TObject);
begin
inherited;
NumPeriode.Value := cNumPeriode.ItemIndex + 1 ;
end;

procedure TFAssistOLE.FlagExtClick(Sender: TObject);
var i : Integer ;
    C : TControl ;
begin
inherited;
if LoadingFormula then exit ;
for i:=0 to Plan.ControlCount-1 do
   begin
   C := TControl(Plan.Controls[i]) ;
   if C is TLabel then continue ;
   if (Copy(C.Name,1,7) = 'FlagExt') then continue ;
   if (Copy(C.Name,1,6) = 'RefCel')  then begin C.Enabled := TCheckBox(Sender).Checked ; continue ; end ;
   C.Enabled := not TCheckBox(Sender).Checked ;
   end ;
if GetPageName = 'ChoixRec' then
   begin
   if FlagExtRec.Checked then lRub.Caption := '' ;
   if not FlagExtRec.Checked then RefCelRec.Text := '' ;
   end ;
if GetPageName = 'ChoixRec2' then
   begin
   if FlagExtRec2.Checked then lRub.Caption := '' ;
   if not FlagExtRec2.Checked then RefCelRec2.Text := '' ;
   end ;
if GetPageName = 'PSQL' then
   begin
   if FlagExtSQL.Checked then lRub.Caption := '' ;
   if not FlagExtSQL.Checked then RefCelSQL.Text := '' ;
   end ;
if GetPageName = 'ChoixBalSit' then
   begin
   if FlagExtBalSit.Checked then lRub.Caption := '' ;
   if not FlagExtBalSit.Checked then RefCelBalSit.Text := '' ;
   end ;
end;

procedure TFAssistOLE.RefCelEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[27] ;
end;

procedure TFAssistOLE.FlagExtEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[28] ;
end;

procedure TFAssistOLE.RefCelRecChange(Sender: TObject);
begin
inherited;
lRub.Caption := MsgEF.Mess[29] + ' ' + RefCelRec.Text ;
end;

procedure TFAssistOLE.rSourceClick(Sender: TObject);
begin
inherited ;
GC2.Visible:=FALSE ;
rBudget.Visible := IsBudgetaire ;
lBudget.Visible := rBudget.Visible ;
gAxe.Visible := IsSection Or IsSection2;
if (gAxe.Visible) and (cAxe.ItemIndex=-1) then
  begin cAxe.ItemIndex := 0 ; cAxeChange(nil) ; end ;
Case rSource.ItemIndex of
   0 : GridFields.Hint := 'RB' ;
   1 : GridFields.Hint := 'RB' ;
   2 : GridFields.Hint := 'G' ;
   3 : GridFields.Hint := 'T' ;
   4 : GridFields.Hint := 'S' ;
   5 : GridFields.Hint := 'BG' ;
   6 : GridFields.Hint := 'BS' ;
   7 : GridFields.Hint := 'J' ;
   8 : GridFields.Hint := 'BJ' ;
   9 : GridFields.Hint := 'SO' ;
   end ;
lSource.Caption := MsgEF.Mess[43 + rSource.ItemIndex] ;
if IsSociete then lRub.Caption := MsgEF.Mess[30] ;
rBudgetClick(nil) ;
GSens.Visible:=RSource.itemIndex=2 ;
LSens.Visible:=RSource.itemIndex=2 ;
CSens.Visible:=RSource.itemIndex=2 ;
GBalSit.Visible:=RSource.ItemIndex in [0,2] ;
If Not GBalSit.Visible Then CCBalSit.Checked:=FALSE ;
If CSens.Visible And (CSens.ItemIndex<0) Then CSens.Value:='SM' ;
If rCumul.Checked Then
  BEGIN
// BPY le 18/12/2003 => Fiche 13086 active cette coche quelque soit la serie !
  {if not(V_PGI.LaSerie=S3) then} GC2.Visible:=TRUE ;
// Fin BPY
  If RSource.ItemIndex in [7..9] Then
    BEGIN
    RSource.ItemIndex:=0 ;
    END ;
  If RSource.ItemIndex in [10..12] Then
    BEGIN
// BPY le 18/12/2003 => Fiche 13086 active cette coche quelque soit la serie !
    {if not(V_PGI.LaSerie=S3) then} GC2.Checked:=TRUE ;
// Fin BPY
    END ;
// BPY le 18/12/2003 => Fiche 13086 active cette coche quelque soit la serie !
  If IsBalSit Then {if not(V_PGI.LaSerie=S3) then} GC2.Checked:=TRUE ;
// Fin BPY
  END ;
  { FQ 15692 - CA - 22/05/2005 - Pour forcer le rechargement de l'onglet avancé }
  OkAvances(TRUE);
end ;

procedure TFAssistOLE.PChange(Sender: TObject);
var titre     : Boolean ;
    s         : string ;
    TobOle    : TOB;
    QLe       : TQuery;
begin
inherited;
titre := (GetPage = ChoixFormule) or (GetPage = ChoixSource) ;
lSource.Visible := not titre ;
lRub.Visible := not titre ;
TitleLine.Visible := not titre ;

if GetPage = ChoixFormule then
   begin
   bPrecedent.Enabled := False ;
   end ;

if GetPage = ChoixSource then
   begin
   if cAxe.ItemIndex=-1 then
      begin cAxe.ItemIndex:=SauveAxe ; cAxeChange(nil) ; end ;
   end ;

if GetPage = ChoixChamp then
   begin
    TobOle := TOB.Create('', nil, -1);
    QLe := OpenSQL('SELECT DH_NOMCHAMP Champ,DH_LIBELLE Libelle,DH_TYPECHAMP TYPECHAMP FROM DECHAMPS WHERE DH_CONTROLE LIKE "%L%" AND DH_PREFIXE = "' + GridFields.Hint + '" ORDER BY DH_NOMCHAMP',TRUE);
    TobOle.LoadDetailDB('DECHAMPS', '', '', QLe, TRUE, FALSE);
    Ferme(QLe);
    GridFields.VidePile(False);
    TobOle.PutGridDetail(GridFields,False,False,'Champ;Libelle;TYPECHAMP');
    TobOle.free;

   bFin.Enabled := True ;
(*AVOIR   if InitialField <> '' then
      begin
      QFields.Locate('DH_NOMCHAMP',InitialField,[]) ;
      InitialField := '' ;
      end ;
*)
   end ;

if GetPage = ChoixRec2 then
     ChangementCode (GridRub2, FindCode2.Text, FindLib2.Text);

if GetPage = ChoixRec then
     ChangementCode (THGRID, FindCode.Text, FindLib.Text);

if GetPage = Periode then
   begin
   if rConstante.Checked then
      begin
      cPeriode.Enabled := False ;
      cPeriode.ItemIndex := 1 ;
      lDetailPeriode.Visible := False ;
      cDetailPeriode.Visible := False ;
      end
   else
      begin
      cPeriode.Enabled := True ;
      lDetailPeriode.Visible := True ;
      cDetailPeriode.Visible := True ;
      end ;
   cPeriodeChange(nil) ;
   CheckExoOKClick(nil) ;
   end ;

if GetPage = Etablissement then
   begin
   s := cEtablissement.Value ;
   cEtablissement.Vide := not rConstante.Checked ;
   cEtablissement.Reload ;
   cEtablissement.Value := s ;
   if cEtablissement.ItemIndex = -1 then cEtablissement.ItemIndex := 0 ;
   end ;

if GetPage = PSQL then
  BEGIN
  BFin.Enabled:=TRUE ;
  END ;

if GetPage = PAvances then
  BEGIN
  BFin.Enabled:=TRUE ;
  If Z_C1.Items.Count<=0 Then OkAvances(TRUE) ;
  END ;

if GetPage = Resultat then
   begin
   ExpliqueFormule ;
   WriteFormule ;
   pFormule.Caption := sFormule ;
   bSuivant.Enabled := False ;
   bFin.Default := True ;
   lAide.Caption := MsgEF.Mess[54] ;
   end ;
end ;

procedure TFAssistOLE.rSourceEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[31] ;
end;

procedure TFAssistOLE.cAxeChange(Sender: TObject);
begin
SauveAxe := cAxe.ItemIndex ;
AxeCourant := cAxe.Value ;
end;

procedure TFAssistOLE.cAxeEnter(Sender: TObject);
begin
inherited;
lAide.Caption := MsgEF.Mess[32] ;
end;

procedure TFAssistOLE.GridRubDblClick(Sender: TObject);
begin
inherited;
if (rCumul.Checked) and (bFin.Enabled) then bFinClick(nil) ;
end;

procedure TFAssistOLE.rCumulClick(Sender: TObject);
begin
// BPY le 18/12/2003 => Fiche 13086 active cette coche quelque soit la serie !
{if not(V_PGI.LaSerie=S3) then} GC2.Visible:=TRUE ;
// Fin BPY
MULTIDOSSIER.Visible := True ;
Regroupement.visible := true ;
ChangeSourceItems(13) ;
DisplayStep ;
bVal.Visible := False ; bLib.Visible := False ;
end;

procedure TFAssistOLE.rChampClick(Sender: TObject);
begin
GC2.Visible:=FALSE ;
MULTIDOSSIER.Visible := False;
Regroupement.visible := False;
ChangeSourceItems(10) ;
DisplayStep ;
bVal.Visible := False ; bLib.Visible := False ;
end;

procedure TFAssistOLE.FormCreate(Sender: TObject);
begin
inherited;
InitialFormula := '' ;
InitialRec := '' ;
InitialRec2 := '' ;
InitialField := '' ;
end;

procedure TFAssistOLE.GridFieldsDblClick(Sender: TObject);
begin
inherited;
if bFin.Enabled then bFinClick(nil) ;
end;

procedure TFAssistOLE.rConstanteClick(Sender: TObject);
begin
MULTIDOSSIER.Visible := False;
Regroupement.visible := False;
GC2.Visible:=FALSE ;
lSource.Caption := MsgEF.Mess[52] ;
DisplayStep ;
bVal.Visible := True ; bLib.Visible := True ;
bVal.Down := True ;
if rConstante.Checked then
   begin
   if Length(InitialRec) > 3 then InitialRec := Copy(InitialRec,1,3) ;
   cPeriode.Enabled := False ;
   cPeriode.ItemIndex := 1 ;
   lDetailPeriode.Visible := False ;
   cDetailPeriode.Visible := False ;
   end ;
end;

procedure TFAssistOLE.NumPeriodeExit(Sender: TObject);
begin
if NumPeriode.Value > NumPeriode.Max then NumPeriode.Value := 1 ;
end;

procedure TFAssistOLE.bValClick(Sender: TObject);
begin
DisplayStep ;
end;

procedure TFAssistOLE.bLibClick(Sender: TObject);
begin
DisplayStep ;
end;

procedure TFAssistOLE.rBudgetClick(Sender: TObject);
begin
if IsBudgete then // Budgeté
   begin
   FlagExtEcr.Checked := False ;
   cNatureEcr.ItemIndex := 0 ;
   cIntegAN.ItemIndex := 1 ;
   end ;
end;

procedure TFAssistOLE.cCptVariantEnter(Sender: TObject);
begin
  inherited;
lAide.Caption := MsgEF.Mess[58] ;

end;

procedure TFAssistOLE.NumPeriodeChange(Sender: TObject);
begin
  inherited;
If NumPeriode1.Value<NumPeriode.Value Then NumPeriode1.Value:=NumPeriode.Value
end;

procedure TFAssistOLE.NumPeriode1Exit(Sender: TObject);
begin
  inherited;
if NumPeriode1.Value > NumPeriode1.Max then NumPeriode1.Value := 1 ;
end;

procedure TFAssistOLE.FindCode2Change(Sender: TObject);
begin
inherited;
If (rSource.ItemIndex in [10..12])=FALSE Then Exit ;
ChangementCode (GridRub2, FindCode2.Text, FindLib2.Text)
end;

procedure TFAssistOLE.FindLib2Change(Sender: TObject);
begin
inherited;
If (rSource.ItemIndex in [10..12])=FALSE Then Exit ;
ChangementCode (GridRub2, FindCode2.Text, FindLib2.Text)
end;

procedure TFAssistOLE.RSQLClick(Sender: TObject);
begin
  inherited;
MULTIDOSSIER.Visible := False;
Regroupement.visible := False;
GC2.Visible:=FALSE ;
//ChangeSourceItems(10) ;
DisplayStep ;
bVal.Visible := False ; bLib.Visible := False ;
end;

Procedure AddZC(ZC : THValCombobox ; TDE : TDEChamp ; OnLib : Boolean) ;
BEGIN
If OnLib Then ZC.Items.Add(TDE.Libelle) Else ZC.Items.Add(TDE.Nom) ;
ZC.Values.Add(TDE.Nom) ;
END ;

procedure TFAssistOLE.ChargeAvances(OnLib : Boolean) ;
Var i,j,k : Integer ;
    SST : Array[1..6] Of String ;
    TC : tControl ;
    ZC : THValComboBox ;
    St : String ;
    av : tav ;
BEGIN
St:='' ; av:=avNone ;
If rEcr.ItemIndex=0 Then BEGIN St:='E' ; av:=avEcr ; END Else
 If rEcr.ItemIndex=1 Then BEGIN St:='Y' ; av:=avAna ; END ;
If CCBalSit.Checked Then av:=avCpt ;
If av=avNone Then Exit ;
i:=PrefixeToNum(St) ;
For k:=1 To 6 Do
  BEGIN
  SST[k]:='' ;
  TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
  If TC<>NIL Then
    BEGIN
    ZC:=THValComboBox(TC) ; If ZC.ItemIndex>=0 Then SST[k]:=ZC.Value ; ZC.Items.Clear ; ZC.Values.Clear ;
    END ;
  END ;
If av in [avEcr,avAna] Then
  BEGIN
  If St<>'' Then
    BEGIN
{$IFDEF EAGLCLIENT}
        ChargeDeChamps(i, St);
{$ENDIF EAGLCLIENT}
    for j:=1 to High(V_PGI.DeChamps[i]) do
      if V_PGI.DEChamps[i,j].Nom<>'' then
        BEGIN
        For k:=1 To 6 Do
          BEGIN
          TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
          If TC<>NIL Then
            BEGIN
            ZC:=THValComboBox(TC) ;
            AddZC(ZC,V_PGI.DEChamps[i,j],OnLib) ;
            END ;
          END ;
        END ;
     END ;
  END ;
St:='G' ;
If av in [avCpt,avEcr,avAna] Then
  BEGIN
  i:=PrefixeToNum(St) ;
{$IFDEF EAGLCLIENT}
    ChargeDeChamps(i, St);
{$ENDIF EAGLCLIENT}

  for j:=1 to High(V_PGI.DeChamps[i]) do
    if V_PGI.DEChamps[i,j].Nom<>'' then
      BEGIN
      For k:=1 To 6 Do
        BEGIN
        TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
        If TC<>NIL Then
          BEGIN
          ZC:=THValComboBox(TC) ;
          AddZC(ZC,V_PGI.DEChamps[i,j],OnLib) ;
          END ;
        END ;
      END ;
    END ;
If av in [avEcr] Then
  BEGIN
  St:='T' ;
  i:=PrefixeToNum(St) ;
{$IFDEF EAGLCLIENT}
    ChargeDeChamps(i, St);
{$ENDIF EAGLCLIENT}
  for j:=1 to High(V_PGI.DeChamps[i]) do
    if V_PGI.DEChamps[i,j].Nom<>'' then
      BEGIN
      For k:=1 To 6 Do
        BEGIN
        TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
        If TC<>NIL Then
          BEGIN
          ZC:=THValComboBox(TC) ;
          AddZC(ZC,V_PGI.DEChamps[i,j],OnLib) ;
          END ;
        END ;
      END ;
  END ;
If av in [avAna] Then
  BEGIN
  St:='S' ;
  i:=PrefixeToNum(St) ;
{$IFDEF EAGLCLIENT}
    ChargeDeChamps(i, St);
{$ENDIF EAGLCLIENT}
  for j:=1 to High(V_PGI.DeChamps[i]) do
    if V_PGI.DEChamps[i,j].Nom<>'' then
      BEGIN
      For k:=1 To 6 Do
        BEGIN
        TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
        If TC<>NIL Then
          BEGIN
          ZC:=THValComboBox(TC) ;
          AddZC(ZC,V_PGI.DEChamps[i,j],OnLib) ;
          END ;
        END ;
      END ;
  END ;
For k:=1 To 6 Do
  BEGIN
  TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
  If TC<>NIL Then
    BEGIN
    ZC:=THValComboBox(TC) ;
    If SST[k]<>'' Then ZC.Value:=SST[k] ;
    END ;
  END ;
END ;

procedure TFAssistOLE.CBLibClick(Sender: TObject);
begin
  inherited;
ChargeAvances(CBLib.Checked) ;
end;

procedure TFAssistOLE.bEffaceAvanceClick(Sender: TObject);
begin
  inherited;
Z_C1.ItemIndex:=-1 ;ZO1.ItemIndex:=-1 ; ZV1.text:='' ; ZG1.ItemIndex:=-1 ;
Z_C2.ItemIndex:=-1 ;ZO2.ItemIndex:=-1 ; ZV2.text:='' ; ZG2.ItemIndex:=-1 ;
Z_C3.ItemIndex:=-1 ;ZO3.ItemIndex:=-1 ; ZV3.text:='' ; ZG3.ItemIndex:=-1 ;
Z_C4.ItemIndex:=-1 ;ZO4.ItemIndex:=-1 ; ZV4.text:='' ; ZG4.ItemIndex:=-1 ;
Z_C5.ItemIndex:=-1 ;ZO5.ItemIndex:=-1 ; ZV5.text:='' ; ZG5.ItemIndex:=-1 ;
Z_C6.ItemIndex:=-1 ;ZO6.ItemIndex:=-1 ; ZV6.text:='' ;
end;

Function TransCompQR(Ch,Comp,Val : String) : String ;
Var St,typ,Val1,Val2 : String ;
BEGIN
if (Ch='') or (Comp='') or (Val='') then BEGIN result:='' ; Exit ; END ;
if (Comp='I') or  (Comp='J') then St:='' else
if (Comp='E') or (Comp='G') then St:=' BETWEEN ' else
if ((Comp='C') or (Comp='L')) then St:=' LIKE ' else
if ((Comp='D') or (Comp='M')) then St:=' NOT LIKE ' else St:=Comp ;

typ:=ChampToType(Ch) ;
if (Comp='E') or (Comp='G') then
   BEGIN
   Val:=FindetReplace(Val,':',';',TRUE) ;
   Val:=FindetReplace(Val,' et ',';',TRUE) ;
   Val:=FindetReplace(Val,' ET ',';',TRUE) ;
   Val:=FindetReplace(Val,' Et ',';',TRUE) ;
   Val1:=ReadTokenSt(Val) ; Val2:=ReadTokenSt(Val) ;
   if Val2='' then Val2:=Val1 ;
   if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val1))+' AND '+IntToStr(ValeurI(Val2)) else
   if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val1))+' AND '+StrfPoint(Valeur(Val2)) else
   if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val1))+'" AND "'+USDATETIME(StrToDate(Val2))+'"' else
      BEGIN
      St:=St+'"'+Val1+'" AND "'+Val2+'"' ;
      END ;
   if Comp='G' then St:=' NOT '+St ;
   END else
if (Comp='I') or (Comp='J') then
   BEGIN
   Val:=FindetReplace(Val,' ou ',';',TRUE) ;
   Val:=FindetReplace(Val,' OU ',';',TRUE) ;
   Val:=FindetReplace(Val,' Ou ',';',TRUE) ;
   Val2:='' ;
   While Val<>'' do
      BEGIN
      Val1:=ReadTokenSt(Val) ;
      if Val1<>'' then
         BEGIN
         if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val1))+', ' else
         if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val1))+' ,' else
         if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val1))+'", ' else
                    St:=St+'"'+Val1+'", ' ;
         Val2:='Ok' ;
         END ;
      END ;
   if Val2<>'' then
      BEGIN
      system.Delete(St,Length(St)-1,2) ;
      St:=' IN ('+St+')' ;
      if Comp='J' then St:=' NOT '+St ;
      END
   END else
   BEGIN
   if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val)) else
   if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val)) else
   if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val))+'"' else
      BEGIN
      St:=St+'"' ;
      if ((Comp='L') or (Comp='M')) then St:=St+'%' ;
      St:=St+CheckdblQuote(Val) ;
      if ((Comp='C') or (Comp='L') or (Comp='D') or (Comp='M')) then St:=St+'%' ;
      St:=St+'"' ;
      END ;
   END ;
Result:=Ch+St ;
Result:=FindetReplace(Result,'"','''',TRUE) ;
END ;

Function TFAssistOLE.FaitSQLSupp : String ;
Var StComp : Array[1..6] Of String ;
    G : Array[1..6] Of Boolean ;
    i : Integer ;
    Z_C,ZO : THvalComboBox ;
    ZV : tEdit ;
    ZG : TComboBox ;
    OkAvance : Boolean ;
    St : String ;
BEGIN
Result:='' ;
Fillchar(StComp,SizeOf(StComp),#0) ; Fillchar(G,SizeOf(G),#0) ;
OkAvance:=FALSE ;
For i:=1 To 6 Do
  BEGIN
  Z_C:=THvalComboBox(FindComponent('Z_C'+IntToStr(i))) ;
  ZO:=THvalComboBox(FindComponent('ZO'+IntToStr(i))) ;
  ZV:=TEdit(FindComponent('ZV'+IntToStr(i))) ;
  If i<6 Then ZG:=THvalComboBox(FindComponent('ZG'+IntToStr(i))) Else ZG:=Nil ;
  If (Z_C<>NIL) And (ZO<>NIL) And (ZV<>NIL) Then
    BEGIN
    StComp[i]:=TransCompQR(Z_C.Value,ZO.Value,ZV.Text) ;
    If Trim(StComp[i])<>'' Then OkAvance:=TRUE ;
    END ;
  If ZG<>NIL Then G[i]:=(ZG.ItemIndex=1) ;
  END ;
If Not OkAvance Then Exit ;
St:='' ;
If ((StComp[1]<>'') and (StComp[2]='') and (StComp[3]='') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if St<>'' then St:='('+St+')' ;
      St:=St+' OR  ('+StComp[1]+')' ;
      END else St:=St+' AND ('+StComp[1]+')' ;
   END else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]='') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+')'
           else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+')' ;
   END else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+')'
              else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+'))' ;
      END else
      BEGIN
      if G[2] then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+'))'
              else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+')' ;
      END ;
   END Else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]<>'') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+')'
                Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+'))' ;
        END Else
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' AND '+StComp[4]+'))' ;
        END ;
      END else
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+'))' ;
        END Else
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' AND '+StComp[4]+')' ;
        END ;
      END ;
   END Else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]<>'') and (StComp[5]<>'')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')'
                  Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+'))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+'))' ;
          END ;
        END Else
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+')))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+'))' ;
          END ;
        END ;
      END else
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+')))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+')))' ;
          END ;
        END Else
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+'))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+')' ;
          END ;
        END ;
      END ;
   END ;

If Trim(St)<>'' Then Delete(St,1,5) ;
Result:=St ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/04/2005
Modifié le ... :   /  /    
Description .. : Accès à l'ongket "Avancé"
Suite ........ : FQ 15692 - CA - 22/04/2005 - Ecritures analytiques pour les
Suite ........ : cas Sections et Généraux/Sections
Mots clefs ... :
*****************************************************************}
Function TFAssistOLE.OkAvances(Recharge : Boolean ; ForceIndex : Integer = -1) : tav ;
BEGIN
Result:=avNone ; rEcr.Enabled:=FALSE ;
If Not GC2.Checked Then Exit ;
If Not rCumul.Checked Then Exit ;
If rSource.ItemIndex in [2,3,10] Then BEGIN If ccBalSit.Checked Then Result:=avCpt Else Result:=avEcr ; If Recharge Then rEcr.ItemIndex:=0 ; END ;
If rSource.ItemIndex in [4,11] Then BEGIN Result:=avAna ;
If Recharge Then
  { FQ 15692 }
  // rEcr.ItemIndex:=0 ;
  rEcr.ItemIndex:=1;
END ;
If rSource.ItemIndex in [0] Then BEGIN If ccBalSit.Checked Then Result:=avCpt Else BEGIN Result:=avEcr ; rEcr.Enabled:=TRUE ; END ;If Recharge Then rEcr.ItemIndex:=0 ; END ;
If (rSource.ItemIndex in [1]) And (Not IsBudgete)Then BEGIN Result:=avEcr ; rEcr.Enabled:=TRUE ; If Recharge Then rEcr.ItemIndex:=0 ; END ;
END ;

procedure TFAssistOLE.rEcrClick(Sender: TObject);
Var GoCharge : Boolean ;
    i : Integer ;
begin
  inherited;
GoCharge:=FALSE ;
If Z_C1.ITems.Count>0 Then
  BEGIN
  If (Copy(Z_C1.Values[0],1,2)='E_') And (rEcr.ItemIndex=1) Then GoCharge:=TRUE ;
  If (Copy(Z_C1.Values[0],1,2)='Y_') And (rEcr.ItemIndex=0) Then GoCharge:=TRUE ;
  END Else GoCharge:=TRUE ;
If GoCharge Then
  BEGIN
  For i:=1 To 6 Do
    BEGIN
    THvalComboBox(FindComponent('Z_C'+IntToStr(I))).Items.Clear ;
    THvalComboBox(FindComponent('Z_C'+IntToStr(I))).Values.Clear ;
    END ;
  ChargeAvances(CBLib.Checked) ;
  END ;
end;

procedure TFAssistOLE.CCBalSitClick(Sender: TObject);
begin
  inherited;
If CCBalSit.Checked Then GC2.Checked:=TRUE ;
end;

procedure TFAssistOLE.GridFieldsClick(Sender: TObject);
begin
  inherited;
   cLibelle.Enabled := (GridFields.Cells[2, GridFields.Row] = 'COMBO') ;
   if not cLibelle.Enabled then cLibelle.Checked := False ;
   bFin.Enabled := True ;
end;

procedure TFAssistOLE.THGRIDClick(Sender: TObject);
begin
  inherited;
   if IsSociete then exit ; // société
   if FlagExtRec.Checked then exit ;
   lRub.Caption := THGRID.Cells[0, THGRID.Row] + ' - ' + THGRID.Cells[1, THGRID.Row];
   if not rChamp.Checked then bFin.Enabled := True ;
end;

procedure TFAssistOLE.GridRub2Click(Sender: TObject);
begin
  inherited;
   if IsSociete then exit ; // société
   if FlagExtRec2.Checked then exit ;
   lRub.Caption := GridRub2.Cells[0, GridRub2.Row] + ' - ' + GridRub2.Cells[1, THGRID.Row];
   if not rChamp.Checked then bFin.Enabled := True ;
end;

procedure TFAssistOLE.FindCodeChange(Sender: TObject);
begin
  inherited;
  ChangementCode (THGRID, FindCode.Text, FindLib.Text);
end;

end.
