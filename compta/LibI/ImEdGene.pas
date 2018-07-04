unit ImEdGene;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QRS1, Db, HMemTb, ExtCtrls, DBTables, Hqry, Menus, HSysMenu, hmsgbox,
  Grids, DBGrids, HDB, HQuickrp, StdCtrls, Hctrls, HTB97, Buttons, ComCtrls,
  HPanel,UiUtil,Hent1, Mask, ParamDat, ImEnt, EdtEtat,LookUp,
  HPdfviewer,UTOB,ImPlan, ImPlanInfo, HStatus, (*QR,*) EdtQR, Spin (*, CalcOLE*),
  uEdtComp, LicUtil, HRichEdt, HRichOLE
  {$IFNDEF SERIE1}
  ,Ent1
  {$ENDIF}
;

type
  TFImEditGen = class(TFQRS1)
    PZLibre: TTabSheet;
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
    HLabel1: THLabel;
    HLabel4: THLabel;
    HLabel2: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    I_LIEUGEO: THValComboBox;
    I_COMPTEREF: THCritMaskEdit;
    I_COMPTEREF_: THCritMaskEdit;
    I_NATUREIMMO: THValComboBox;
    _I_DATEPIECEA: THCritMaskEdit;
    _I_DATEPIECEA_: THCritMaskEdit;
    XX_WHERE: TEdit;
    I_DATEPIECEA: THCritMaskEdit;
    I_DATEPIECEA_: THCritMaskEdit;
    PzLibreS1: TTabSheet;
    TT_TABLELIBREIMMO1: THLabel;
    TT_TABLELIBREIMMO2: THLabel;
    TT_TABLELIBREIMMO3: THLabel;
    TABLELIBRE3: THValComboBox;
    TABLELIBRE2: THValComboBox;
    TABLELIBRE1: THValComboBox;
    PComplement: TTabSheet;
    LabelDateCession: THLabel;
    ARRETDOTATION: THCritMaskEdit;
    tI_ETABLISSEMENT: THLabel;
    I_ETABLISSEMENT: THValComboBox;
    LArretDotation: THLabel;
    _I_DATECESSION: THCritMaskEdit;
    _I_DATECESSION_: THCritMaskEdit;
    LabelAuDateCession: THLabel;
    I_QUALIFIMMO: THValComboBox;
    HLabel3: THLabel;
    SYNTH: TCheckBox;
    HLabel8: THLabel;
    I_IMMO: THCritMaskEdit;
    I_IMMO_: THCritMaskEdit;
    HLabel9: THLabel;
    function  RecupWhereSQL : hstring ;override;
    procedure BValiderClick(Sender: TObject); override ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DATEKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure IMMOElipsisClick(Sender: TObject);
    procedure I_COMPTEREFElipsisClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure FEtatChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TABLELIBRE1Change(Sender: TObject);
    procedure I_NATUREIMMOChange(Sender: TObject);
  private
    { Déclarations privées }
    bArretDotation : boolean;
    fTypeEdition : string;
    procedure AssignHelpContext;
    procedure UpdateZones;
    function ControleZones: boolean;
  public
    { Déclarations publiques }
  end;

procedure LanceEditionImmo (TypeEdition : string);

implementation

uses
{$IFDEF SERIE1}
{$ELSE}
ImOle,
{$ENDIF}
ImEdCalc ;


{$R *.DFM}

procedure LanceEditionImmo (TypeEdition : string);
var   FImEditGen: TFImEditGen; PP : THPanel;
      Nature : string;
begin
{$IFNDEF IMP}
{$IFDEF SERIE1}
{$ELSE}
ProcZoomEdt:=ImZoomEdtEtatImmo ;
ProcCalcEdt:=CalcOLEEtatImmo ;
{$ENDIF}
{$ENDIF}
FImEditGen := TFImEditGen.Create (Application);
PP:=FindInsidePanel ;
FImEditGen.fTypeEdition:=TypeEdition ;

if (TypeEdition = 'DCB') or (TypeEdition = 'ECB') then Nature := 'AMC'
else if (TypeEdition = 'DDE') or (TypeEdition = 'DEC')
      or (TypeEdition = 'DFI') or (TypeEdition = 'DOT') then Nature := 'AMD'
else if (TypeEdition = 'ACS') or (TypeEdition = 'MUT')
      or (TypeEdition = 'PMV') or (TypeEdition = 'STS') then Nature := 'AMO'
else  if (TypeEdition = 'PRE') or (TypeEdition = 'PRF') then Nature := 'AMP'
else Nature := 'AML';

{$IFDEF SERIE1}
//FImEditGen.InitQR('E','IMM',TypeEdition,false,true) ;
FImEditGen.ProposePasEtat:=true ; //YCP sinon l'initialisation ne se fait pas - MVG 12/07/2006
FImEditGen.InitQR('E',Nature,TypeEdition,false,true) ;
{$ELSE}
FImEditGen.InitQR('E',Nature,TypeEdition,True,true) ;
{$ENDIF}
if PP=Nil then
  begin
  try
    FImEditGen.WindowState := wsMaximized;
    FImEditGen.ShowModal ;
    finally
    FImEditGen.Free ;
    end ;
  end
else
  begin
  InitInside(FImEditGen,PP) ;
  FImEditGen.Show ;
  end;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFImEditGen.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  {$IFDEF SERIE1}
  {$ELSE}
  ProcZoomEdt:=ImZoomEdtEtat ;
  ProcCalcEdt:=ImCalcOLEEtat ;
  {$ENDIF}
  if isInside(Self) then Action:=caFree ;
end;

procedure TFImEditGen.DATEKeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
  ParamDate (Self, Sender, Key);
end;

procedure TFImEditGen.FormShow(Sender: TObject);
begin
  inherited;
  bEuro.Visible:=false ;
  bProg.Visible:=false ;
  {$IFDEF SERIE1}
  I_ETABLISSEMENT.Visible := False;
  tI_ETABLISSEMENT.Visible := False;
  {$ENDIF}
  {$IFNDEF SERIE1}
  // Paramétrage des états possible si entrée avec le mot de passe du jour
  //bParamEtat.Visible := (not (ctxPCL in V_PGI.PGIContexte)) or (V_PGI.PassWord = CryptageSt(DayPass(Date)));
  bParamEtat.Visible := True;
  {$ENDIF}
  bEuro.visible:=false ;
  bProg.visible:=false ;
  Caption := FEtat.Text;
  UpdateCaption(self);
  AssignHelpContext;
  _I_DATEPIECEA.Text:=StDate1900;
  _I_DATEPIECEA_.Text:=StDate2099;
  _I_DATECESSION.Text := StDate1900;
  _I_DATECESSION_.Text := StDate2099;
  I_COMPTEREF.Text :=VHImmo^.CpteImmoInf;
  I_COMPTEREF_.Text:=VHImmo^.CpteLocSup;
  I_DATEPIECEA.Text:=StDate1900;
  I_DATEPIECEA_.Text:=StDate2099;
  I_QUALIFIMMO.Value:='R' ;
  UpdateZones;
  {$IFNDEF SERIE1}
  PositionneEtabUser(I_ETABLISSEMENT);
  {$ENDIF}
end;

procedure TFImEditGen.IMMOElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpList (TControl (Sender),'Immobilisation','IMMO','I_IMMO','I_LIBELLE','','I_IMMO',True,0) ;
end;

procedure TFImEditGen.I_COMPTEREFElipsisClick(Sender: TObject);
var CpteInf,CpteSup : string;
    sWhere : string;
begin
  inherited;
  if I_NATUREIMMO.Value='CB' then begin CpteSup := VHImmo^.CpteCBSup; CpteInf := VHImmo^.CpteCBInf;end
  else if I_NATUREIMMO.Value='LOC' then begin CpteSup := VHImmo^.CpteLocSup; CpteInf := VHImmo^.CpteLocInf; end
  else if I_NATUREIMMO.Value='PRO' then begin CpteSup := VHImmo^.CpteImmoSup;CpteInf := VHImmo^.CpteImmoInf;end
  else if I_NATUREIMMO.Value='FI'  then begin CpteSup := VHImmo^.CpteFinSup; CpteInf := VHImmo^.CpteFinInf; end
  else begin CpteSup := VHImmo^.CpteLocSup; CpteInf :=VHImmo^.CpteImmoInf;end;
  sWhere:='G_GENERAL<="'+CpteSup+'" AND G_GENERAL>="'+CpteInf+'"';
  LookUpList (TControl (Sender),Msg.Mess[3],'GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,0) ;
end;

function TFImEditGen.RecupWhereSQL : hstring ;
var DateCDeb,DateCFin,DateADeb,DateAFin: string ;
begin
  XX_WHERE.Text := '';
  DateCDeb:=USDateTime(StrToDate(_I_DATECESSION.Text)) ;
  DateCFin:=USDateTime(StrToDate(_I_DATECESSION_.Text)) ;
  DateADeb:=USDateTime(StrToDate(_I_DATEPIECEA.Text)) ;
  DateAFin:=USDateTime(StrToDate(_I_DATEPIECEA_.Text)) ;

  if (fTypeEdition='DOT') or (fTypeEdition='DFI')
  or (fTypeEdition='DEC') or (fTypeEdition='DDE')
  or (fTypeEdition='ACS') or (fTypeEdition='STS')
  then
  begin
    XX_WHERE.Text:='I_ETAT<>"FER" '
                 + ' AND ((IL_TYPEOP LIKE "CE%" AND IL_DATEOP>="'+DateCDeb+'" AND IL_DATEOP<="'+DateCFin+'") '
                 +  'OR (IL_TYPEOP="ACQ" AND I_DATEPIECEA>="'+DateADeb+'" AND I_DATEPIECEA<="'+DateAFin+'")) ';
  end
  else
  begin
    I_DATEPIECEA.Text :=_I_DATEPIECEA.Text;
    I_DATEPIECEA_.Text:=_I_DATEPIECEA_.Text;
  end ;

  if (fTypeEdition='DDE') or (fTypeEdition='DFI') then XX_WHERE.Text:=XX_WHERE.Text+' AND I_METHODEFISC <>""';
  if (fTypeEdition='DCB') then XX_WHERE.Text:='(I_ETAT<>"FER")';
  XX_WHERE.Text := inherited RecupWhereSQL;
  if (fTypeEdition='DCB') or (fTypeEdition='ECB') then XX_WHERE.Text:=FindEtReplace(XX_WHERE.Text,'I_COMPTEREF','I_COMPTELIE',true);
  result:=XX_WHERE.Text;
end;

procedure TFImEditGen.HMTradBeforeTraduc(Sender: TObject);
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

function TFImEditGen.ControleZones : boolean;
begin
  result:=false ;
  if (not IsValidDate(_I_DATEPIECEA.Text)) or (not IsValidDate(_I_DATEPIECEA.Text)) then Msg.Execute(5,Caption,'')
  else if (bArretDotation) then
  begin
    if (not IsValidDate(ARRETDOTATION.Text)) then Msg.Execute(5,Caption,'')
    else if (StrToDate(ARRETDOTATION.Text)<VHImmo^.Encours.Deb) or (StrToDate(ARRETDOTATION.Text)>VHImmo^.Encours.Fin) then
    begin
      Msg.Execute(4,'','');
      ARRETDOTATION.Text := DateToStr(VHImmo^.Encours.Fin);
    end
    else
      result:=true ;
  end
  else
    result:=true;
end;

procedure TFImEditGen.FEtatChange(Sender: TObject);
begin
  inherited;
  AssignHelpContext;
  UpdateZones;
end;

procedure TFImEditGen.UpdateZones;
begin
  fTypeEdition := fCodeEtat;
  bArretDotation :=(fTypeEdition='DOT') or (fTypeEdition='DFI') or (fTypeEdition='DDE') or (fTypeEdition='DEC') or (fTypeEdition='DCB');
  ARRETDOTATION.Visible :=bArretDotation;
  LArretDotation.Visible:=bArretDotation;
  if ARRETDOTATION.Visible then ARRETDOTATION.Text:=DateToStr(VHImmo^.Encours.Fin) ;
  {$IFDEF SERIE1}
  FListe.Visible:=true ; //ycp false ;
  {$ELSE}
  FListe.Visible:=(fTypeEdition <> 'ETI') and (fTypeEdition <> 'FIM');
  {$ENDIF}
  if ftypeEdition='ACS' then
  begin
    _I_DATEPIECEA.Text:=DateToStr(VHImmo^.Encours.Deb) ;
    _I_DATEPIECEA_.Text:=DateToStr(VHImmo^.Encours.Fin) ;
    _I_DATECESSION.Text:=DateToStr(VHImmo^.Encours.Deb) ;
    _I_DATECESSION_.Text:=DateToStr(VHImmo^.Encours.Fin) ;
  end ;

  if bArretDotation then     _I_DATEPIECEA_.Text:=DateToStr(VHImmo^.Encours.Fin) ;
  if (fTypeEdition='DOT') or (fTypeEdition='DEC') or (fTypeEdition='DFI') or (fTypeEdition='DDE') or (fTypeEdition='SIM') then
    I_NATUREIMMO.Value:='PRO'
  else if (fTypeEdition='DCB') or (fTypeEdition='ECB') then
    I_NATUREIMMO.Value:='CB'
  else if (fTypeEdition<>'ACS') and (fTypeEdition<>'STS') then
  begin
    LabelDateCession.visible := false;
    _I_DATECESSION.visible :=false;
    LabelAUDateCession.visible := false;
    _I_DATECESSION_.visible:=false;
  end;

end;

procedure TFImEditGen.BValiderClick(Sender: TObject);
begin
  Hent1.EnableControls(Self,false,true) ;
  if ControleZones then
    begin
    VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;
    inherited;
    end;
  Hent1.EnableControls(Self,true,true) ;
  VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;  
end;

procedure TFImEditGen.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
Case Key of
  VK_F9,VK_F10 : BEGIN  if not ControleZones then exit; END ;
  END ;
  inherited;
end;

procedure TFImEditGen.AssignHelpContext;
begin
{$IFDEF SERIE1}
       if (fCodeEtat='FIM') then HelpContext:=550500
  else if (fCodeEtat='DOT') then HelpContext:=552100
  else if (fCodeEtat='DEC') then HelpContext:=552200
  else if (fCodeEtat='DFI') then HelpContext:=552300
  else if (fCodeEtat='DER') then HelpContext:=552400
  else if (fCodeEtat='INV') then HelpContext:=551200
  else if (fCodeEtat='ACS') then HelpContext:=551100
  else if (fCodeEtat='SIM') then HelpContext:=551300
  else if (fCodeEtat='STS') then HelpContext:=551400
  else if (fCodeEtat='PMV') then HelpContext:=551600
  else if (fCodeEtat='MUT') then HelpContext:=551500
  else if (fCodeEtat='DCB') then HelpContext:=553500
  else if (fCodeEtat='ECB') then HelpContext:=553600
  else if (fCodeEtat='PTP') then HelpContext:=553900
  else if (fCodeEtat='PRE') then HelpContext:=553100
  else if (fCodeEtat='PRF') then HelpContext:=553200
  else if (fCodeEtat='ETI') then HelpContext:=554000 ;
{$ELSE}
  if (fCodeEtat='FIM') then HelpContext := 2505000
  else if (fCodeEtat='DOT') then HelpContext := 2521000
  else if (fCodeEtat='DEC') then HelpContext := 2522000
  else if (fCodeEtat='DFI') then HelpContext := 2523000
  else if (fCodeEtat='DER') then HelpContext := 2524000
  else if (fCodeEtat='INV') then HelpContext := 2512000
  else if (fCodeEtat='ACS') then HelpContext := 2511000
  else if (fCodeEtat='SIM') then HelpContext := 2513000
  else if (fCodeEtat='STS') then HelpContext := 2514000
  else if (fCodeEtat='PMV') then HelpContext := 2516000
  else if (fCodeEtat='MUT') then HelpContext := 2515000
  else if (fCodeEtat='DCB') then HelpContext := 2535000
  else if (fCodeEtat='ECB') then HelpContext := 2536000
  else if (fCodeEtat='PTP') then HelpContext := 2539000
  else if (fCodeEtat='PRE') then HelpContext := 2531000
  else if (fCodeEtat='PRF') then HelpContext := 2532000
  else if (fCodeEtat='ETI') then HelpContext := 2540000;
{$ENDIF}
end;

procedure TFImEditGen.TABLELIBRE1Change(Sender: TObject);
begin
  inherited;
  if Sender=TABLELIBRE1 then I_TABLE0.Text:=TABLELIBRE1.Value
  else if Sender=TABLELIBRE2 then I_TABLE1.Text:=TABLELIBRE2.Value
  else if Sender=TABLELIBRE3 then I_TABLE2.Text:=TABLELIBRE3.Value
end;

procedure TFImEditGen.I_NATUREIMMOChange(Sender: TObject);
begin
  inherited;
  if I_NATUREIMMO.Value='CB' then begin I_COMPTEREF_.Text:=VHImmo^.CpteCBSup; I_COMPTEREF.Text:=VHImmo^.CpteCBInf; end
  else if I_NATUREIMMO.Value='LOC' then begin I_COMPTEREF_.Text:=VHImmo^.CpteLocSup; I_COMPTEREF.Text:=VHImmo^.CpteLocInf;  end
  else if I_NATUREIMMO.Value='PRO' then begin I_COMPTEREF_.Text:=VHImmo^.CpteImmoSup;I_COMPTEREF.Text:=VHImmo^.CpteImmoInf; end
  else if I_NATUREIMMO.Value='FI'  then begin I_COMPTEREF_.Text:=VHImmo^.CpteFinSup; I_COMPTEREF.Text:=VHImmo^.CpteFinInf;  end
  else begin I_COMPTEREF_.Text:=VHImmo^.CpteLocSup; I_COMPTEREF.Text:=VHImmo^.CpteImmoInf; end;
end;

end.


