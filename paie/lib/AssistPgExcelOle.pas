{Modifications
PT1 GGS 03/2007 Liens OLE MultiDossier

}
unit AssistPgExcelOle;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, ComCtrls, ExtCtrls, HPanel, HTB97,

{$IFDEF EAGLCLIENT}
     eFichList,
     UTOB,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables, Mask, Hctrls{$ELSE}uDbxDataSet{$ENDIF},
     {$IFNDEF EAGLSERVER}
     FichList,
     {$ENDIF}
{$ENDIF}
  Hctrls, HFLabel, Mask, Hent1, ParamDat, uTOM;

type
  TFAssistExcelOle = class(TFAssist)
    TBCaract: TTabSheet;
    RBBilansocial: TRadioButton;
    RBBilanSocialSimple: TRadioButton;
    RBCumulPaie: TRadioButton;
    RBEffectif: TRadioButton;
    TBBilanSocial: TTabSheet;
    PBC_ETABLISSEMENT: THMultiValComboBox;
    PBC_INDICATEURBS: THValComboBox;
    PBC_BSPRESENTATION: THValComboBox;
    TPBC_INDICATEURBS: THLabel;
    TPBC_BSPRESENTATION: THLabel;
    PBC_DATEDEBUT: THCritMaskEdit;
    TPBC_DATEDEBUT: THLabel;
    TPBC_DATEFIN: THLabel;
    PBC_DATEFIN: THCritMaskEdit;
    TPBC_ETABLISSEMENT: THLabel;
    TPBC_CATBILAN: THLabel;
    PBC_CATBILAN: THMultiValComboBox;
    TBComplMulti: TTabSheet;
    TMCODESTAT: THLabel;
    MCODESTAT: THMultiValComboBox;
    TMTRAVAILN1: THLabel;
    MTRAVAILN1: THMultiValComboBox;
    TMTRAVAILN2: THLabel;
    MTRAVAILN2: THMultiValComboBox;
    TMTRAVAILN3: THLabel;
    MTRAVAILN3: THMultiValComboBox;
    TMTRAVAILN4: THLabel;
    MTRAVAILN4: THMultiValComboBox;
    TPBC_LIBELLEEMPLOI: THLabel;
    PBC_LIBELLEEMPLOI: THMultiValComboBox;
    PBC_COEFFICIENT: THMultiValComboBox;
    PBC_QUALIFICATION: THMultiValComboBox;
    TPBC_QUALIFICATION: THLabel;
    TPBC_COEFFICIENT: THLabel;
    TMLIBREPCMB1: THLabel;
    MLIBREPCMB1: THMultiValComboBox;
    TMLIBREPCMB2: THLabel;
    MLIBREPCMB2: THMultiValComboBox;
    TMLIBREPCMB3: THLabel;
    MLIBREPCMB3: THMultiValComboBox;
    TMLIBREPCMB4: THLabel;
    MLIBREPCMB4: THMultiValComboBox;
    TBPAIECUMUL: TTabSheet;
    HDATEDEB: THLabel;
    DATEDEB: THCritMaskEdit;
    DATEFIN: THCritMaskEdit;
    HDATEFIN: THLabel;
    HEtab: TLabel;
    HSAL: TLabel;
    SALARIE: THCritMaskEdit;
    TYPEDONNEE: THValComboBox;
    HTYPEDONNEE: TLabel;
    HTYPEVALEUR: TLabel;
    TYPEVALEUR: THValComboBox;
    HRUBRIQUE: TLabel;
    HCATDADS: TLabel;
    HDADSPROF: TLabel;
    TBComplCombo: TTabSheet;
    TPSA_CODESTAT: THLabel;
    TPSA_TRAVAILN1: THLabel;
    TPSA_TRAVAILN2: THLabel;
    TPSA_TRAVAILN3: THLabel;
    TPSA_TRAVAILN4: THLabel;
    TPSA_LIBREPCMB1: THLabel;
    TPSA_LIBREPCMB2: THLabel;
    TPSA_LIBREPCMB3: THLabel;
    TPSA_LIBREPCMB4: THLabel;
    PSA_CODESTAT: THValComboBox;
    PSA_TRAVAILN1: THValComboBox;
    PSA_TRAVAILN2: THValComboBox;
    PSA_TRAVAILN3: THValComboBox;
    PSA_TRAVAILN4: THValComboBox;
    PSA_LIBREPCMB1: THValComboBox;
    PSA_LIBREPCMB2: THValComboBox;
    PSA_LIBREPCMB3: THValComboBox;
    PSA_LIBREPCMB4: THValComboBox;
    Present: TCheckBox;
    TBPaieBSCalInfo: TTabSheet;
    HTYPERENS: TLabel;
    TYPERENS: THValComboBox;
    HTYPEVAL1: TLabel;
    TYPEVAL1: THValComboBox;
    HDATED: THLabel;
    DATED: THCritMaskEdit;
    DATEF: THCritMaskEdit;
    HDATEF: THLabel;
    CAT: THCritMaskEdit;
    HCATDUCS: TLabel;
    VAL1: THCritMaskEdit;
    VAL2: THCritMaskEdit;
    HVAL1: TLabel;
    HVAL2: TLabel;
    RBPaieInfo: TRadioButton;
    TBPaieRendInfo: TTabSheet;
    HTYPETABLE: TLabel;
    TYPETABLE: THValComboBox;
    SALARIE2: THCritMaskEdit;
    HSALARIE2: TLabel;
    Label1: TLabel;
    CHAMP: THValComboBox;
    QUOI: THValComboBox;
    HQUOI: TLabel;
    ETABLISSEMENT: THMultiValComboBox;
    RUBRIQUE: THMultiValComboBox;
    CATDADS: THMultiValComboBox;
    DADSPROF: THMultiValComboBox;
    TBPaieEffectif: TTabSheet;
    MENSUALITE: THCritMaskEdit;
    HLabel2: THLabel;
    Label2: TLabel;
    PETABLISSEMENT: THMultiValComboBox;
    HPSEXE: THLabel;
    PSEXE: THValComboBox;
    TBPaieEffectifCompl: TTabSheet;
    HLIBELLEEMPLOI: THLabel;
    PLIBELLEEMPLOI: THMultiValComboBox;
    HPCOEFFICIENT: THLabel;
    PCOEFFICIENT: THMultiValComboBox;
    HPQUALIFICATION: THLabel;
    PQUALIFICATION: THMultiValComboBox;
    HPDADSCAT: TLabel;
    PDADSCAT: THMultiValComboBox;
    PDADSPROF: THMultiValComboBox;
    HPDADSPROF: TLabel;
    PCATDUCS: THMultiValComboBox;
    HPCATDUCS: TLabel;
    PNIVEAU: THMultiValComboBox;
    HPNIVEAU: THLabel;
    HPINDICE: THLabel;
    PINDICE: THMultiValComboBox;
    PCODEEMPLOI: THMultiValComboBox;
    HPCODEEMPLOI: THLabel;
    PTYPECONTRAT: THValComboBox;
    HPTYPECONTRAT: THLabel;
    HPENTSORT: THLabel;
    PENTSORT: THValComboBox;
    PMOTIFCONT: THValComboBox;
    HPMOTIFCONT: THLabel;
    HPSOMEFF: THLabel;
    PSOMEFF: THValComboBox;
//Début PT1
    TMULTIDOSSIER: TLabel;
    MULTIDOSSIER: THValComboBox;
    Bevel1: TBevel;
    RBChoixMultiSOc: TRadioButton;
    SELECTIONSOCIETE: THMultiValComboBox;
    TSelectionSociete: TLabel;
//Fin PT1
    procedure FormShow(Sender: TObject);
    function  NextPage: TTabSheet; Override ;
    function  PreviousPage : TTabSheet ; Override ;
    procedure PgParamDate(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure PBC_BSPRESENTATIONChange(Sender: TObject);
    procedure TYPEVALEURChange(Sender: TObject);
    procedure SALARIEExit(Sender: TObject);
    procedure TYPERENSChange(Sender: TObject);
    procedure PENTSORTChange(Sender: TObject);
    procedure PTYPECONTRATChange(Sender: TObject);
    procedure MULTIDOSSIERSelect(Sender: TObject);        //PT1
    procedure RBBilansocialClick(Sender: TObject);
    procedure RBChoixMultiSOcClick(Sender: TObject);      //PT1
    procedure RBBilanSocialSimpleClick(Sender: TObject);
    procedure RBCumulPaieClick(Sender: TObject);
    procedure RBEffectifClick(Sender: TObject);
    procedure RBPaieInfoClick(Sender: TObject);
  private
    { Déclarations privées }
    OkZoneLibre : Boolean;
    TabValCal : String;
    Function GenereFormule : Boolean;
    Function GetVerbeOleBilanSocial : boolean;
    Function GetVerbeOlePaieCumul : boolean;
    Function GetVerbeOlePaieBSCalInfo : boolean;
    Function GetVerbeOlePaieRendInfo : boolean;
    Function GetVerbeOlePaieEffectif : Boolean;
    procedure VisibiliteChampLibre(Cont,TCont : TControl; Lib : String );
    procedure SupprimePointVirguleFin ( Var St : string);
  public
    { Déclarations publiques }
    FormuleOle : String;
    Error : integer;
  end;

const
  // libellés des messages d'anomalies
  TexteMessage: array[1..14] of string = (
    {1}'Le format de la date de début est incorrect.',
    {2}'Le format de la date de fin est incorrect.',
    {3}'Vous devez sélectionner une présentation.',
    {4}'Vous devez sélectionner un indicateur.',
    {5}'Vous devez sélectionner un établissement.',
    {6}'Vous devez sélectionner un matricule salarié.',
    {7}'Vous devez sélectionner un type de données.',
    {8}'Vous devez sélectionner un type de valeur.',
    {9}'Vous devez sélectionner un cumul de paie.',
   {10}'Vous devez sélectionner une rubrique de rémunération.',
   {11}'Vous devez sélectionner une rubrique de cotisation.',
   {12}'Vous devez renseigner le type de valeur ou les valeurs.',
   {13}'Vous devez renseigner le champ.',
   {14}'Vous devez renseigner la table.'
    ) ;



function LancePgAssistExcelOle(Sender : Tobject = nil) : string ;

var
  FAssistExcelOle: TFAssistExcelOle;

implementation
uses Entpaie,PgOutils2, OlePaiePgi, UtilPGI;

function LancePgAssistExcelOle(Sender : Tobject = nil) : string;
var    AssistPGExcelOle : TFAssistExcelOle;
Begin

  AssistPGExcelOle := TFAssistExcelOle.Create(Application);
  try
    if Assigned(Sender) and (Sender is TOlePaiePGI) then
    begin
      AssistPGExcelOle.MULTIDOSSIER.Value := (Sender as TOlePaiePGI).RegroupementSociete;
      AssistPGExcelOle.SELECTIONSOCIETE.Value := (Sender as TOlePaiePGI).listebases;
    end;

    AssistPGExcelOle.ShowModal;
    Result:= AssistPGExcelOle.FormuleOle ;

    if Assigned(Sender) and (Sender is TOlePaiePGI) then      
    begin
      (Sender as TOlePaiePGI).RegroupementSociete := AssistPGExcelOle.MULTIDOSSIER.Value;
      (Sender as TOlePaiePGI).listebases := AssistPGExcelOle.SELECTIONSOCIETE.Value;
    end;
  finally
    AssistPGExcelOle.Free;
  end;

End;


Function GetBasesMSWithLib (CodeRegroupement : string = '') : String ;  //PT1
var lQMD     : TQuery ;
    lTSData  : TStringList ;
    lStVal   : String ;
    lStBase  : String ;
begin
  result := '' ;
  if CodeRegroupement = '' then CodeRegroupement := MS_CODEREGROUPEMENT;
  // récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD   := OpenSQL('SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + CodeRegroupement + '"', True ) ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;
  // Récupération 1ère ligne
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;
  // On ne garde que le nom des bases
  while lStVal<>'' do
  begin
    lStBase := ReadTokenSt( lStVal ) ;
    result := result + lStBase + ';';
  end;
  FreeAndNil( lTSData ) ;
end ;
{$R *.DFM}

{ TFAssistExcelOle }

function TFAssistExcelOle.NextPage: TTabSheet;
begin
  Result := Nil;
  BFin.enabled := False;
  if P.ActivePage = TBCaract then
     Begin
     if RBBilansocial.Checked then Result := TBBilanSocial
     else
       if RBCumulPaie.Checked   then  Result := TBPaieCumul
       else
         if RBBilanSocialSimple.Checked then Result := TBPaieBSCalInfo
       else
         if RBPaieInfo.Checked then Result := TBPaieRendInfo
       else
         if RBEffectif.Checked then Result := TBPaieEffectif
     End
  {else
    if ((P.ActivePage = TBBilanSocial) and (OkZoneLibre)) then
       begin
       Result := TBBilanSocialSuite;
       BFin.enabled := True;
       BSuivant.Enabled := False ;
       end   }
  else
    if (((P.ActivePage = TBPaieCumul) OR (P.ActivePage = TBPaieEffectifCompl)) and (OkZoneLibre)) then
      Begin
      Result := TBComplMulti;
      BFin.enabled := True;
      BSuivant.Enabled := False ;
      end
    else
      if (P.ActivePage = TBPaieEffectif) and (OkZoneLibre) then
        Begin
        Result := TBPaieEffectifCompl;
        BFin.enabled := True;
        BSuivant.Enabled := False ;
        end
      else
        Begin
        bSuivant.Enabled := False ;
        BFin.enabled := True;
        Result := Nil;
        End;
end;

function TFAssistExcelOle.PreviousPage: TTabSheet;
begin
Result := Nil;
BSuivant.Enabled := True ;
BFin.enabled := False;
if (P.ActivePage = TBBilanSocial) or (P.ActivePage = TBPaieCumul)
or (P.ActivePage = TBPaieBSCalInfo) or (P.ActivePage = TBPaieRendInfo)
or (P.ActivePage = TBPaieEffectif)  then Result := TBCaract;
//if P.ActivePage = TBBilanSocialSuite then Result := TBBilanSocial;
if (P.ActivePage = TBComplMulti) AND (RBCumulPaie.Checked) then
   Result := TBPaieCumul
else
  if (P.ActivePage = TBComplMulti) AND (RBEffectif.Checked) then
       Result := TBPaieEffectifCompl;

if (P.ActivePage =TBPaieEffectifCompl) then Result := TBPaieEffectif;
end;

procedure TFAssistExcelOle.PgParamDate(Sender: TObject);
var key : char;
begin
  inherited;
  key := '*';
  ParamDate (Self, Sender, Key);
end;

procedure TFAssistExcelOle.FormShow(Sender: TObject);
var
  SavMULTIDOSSIER : String;  //PT1
begin
//PT1 sauve regroupement saisi dans assistant
SavMULTIDOSSIER := MULTIDOSSIER.Value;
inherited;
//PT1 récupère regroupement sauvegardé
MULTIDOSSIER.Value := SavMULTIDOSSIER;
Bfin.enabled := True;
MakeZoomOLE(Handle) ;
TabValCal := '';
TYPEVALEURChange(nil);
PBC_BSPRESENTATIONchange(nil);
FormuleOle := '';
OkZoneLibre := False;

VisibiliteChampLibre(MCODESTAT,TMCODESTAT,Vh_paie.PGLibCodeStat);
VisibiliteChampLibre(PSA_CODESTAT,TPSA_CODESTAT,Vh_paie.PGLibCodeStat);

VisibiliteChampLibre(MTRAVAILN1,TMTRAVAILN1,Vh_paie.PGLibelleOrgStat1);
VisibiliteChampLibre(MTRAVAILN2,TMTRAVAILN2,Vh_paie.PGLibelleOrgStat2);
VisibiliteChampLibre(MTRAVAILN3,TMTRAVAILN3,Vh_paie.PGLibelleOrgStat3);
VisibiliteChampLibre(MTRAVAILN4,TMTRAVAILN4,Vh_paie.PGLibelleOrgStat4);
VisibiliteChampLibre(PSA_TRAVAILN1,TPSA_TRAVAILN1,Vh_paie.PGLibelleOrgStat1);
VisibiliteChampLibre(PSA_TRAVAILN2,TPSA_TRAVAILN2,Vh_paie.PGLibelleOrgStat2);
VisibiliteChampLibre(PSA_TRAVAILN3,TPSA_TRAVAILN3,Vh_paie.PGLibelleOrgStat3);
VisibiliteChampLibre(PSA_TRAVAILN4,TPSA_TRAVAILN4,Vh_paie.PGLibelleOrgStat4);


VisibiliteChampLibre(MLIBREPCMB1,TMLIBREPCMB1,Vh_paie.PgLibCombo1);
VisibiliteChampLibre(MLIBREPCMB2,TMLIBREPCMB2,Vh_paie.PgLibCombo2);
VisibiliteChampLibre(MLIBREPCMB3,TMLIBREPCMB3,Vh_paie.PgLibCombo3);
VisibiliteChampLibre(MLIBREPCMB4,TMLIBREPCMB4,Vh_paie.PgLibCombo4);
VisibiliteChampLibre(PSA_LIBREPCMB1,TPSA_LIBREPCMB1,Vh_paie.PgLibCombo1);
VisibiliteChampLibre(PSA_LIBREPCMB2,TPSA_LIBREPCMB2,Vh_paie.PgLibCombo2);
VisibiliteChampLibre(PSA_LIBREPCMB3,TPSA_LIBREPCMB3,Vh_paie.PgLibCombo3);
VisibiliteChampLibre(PSA_LIBREPCMB4,TPSA_LIBREPCMB4,Vh_paie.PgLibCombo4);

end;


procedure TFAssistExcelOle.VisibiliteChampLibre(Cont, TCont: TControl; Lib : String);
begin
if not Assigned(Cont) then exit;
Cont.Visible := (Lib <> '');
THLAbel(TCont).Visible := (Lib <> '');
THLAbel(TCont).Caption := Lib;
if OkZoneLibre = False then OkZoneLibre := (Lib <> '');
end;


procedure TFAssistExcelOle.bFinClick(Sender: TObject);
begin
  inherited;
If GenereFormule Then ModalResult:=mrOk
else
  Begin
  if (Error <> 0) then PgiBox(TexteMessage[Error] , caption); //+ ComplMessage
  ModalResult:=MrNone;
  End;
end;

function TFAssistExcelOle.GenereFormule: Boolean;
Begin
  Result := False;
  if RBBilansocial.Checked then Result := GetVerbeOleBilanSocial
  else
    if RBCumulPaie.Checked then Result := GetVerbeOlePaieCumul
  else
    if RBBilanSocialSimple.Checked then Result := GetVerbeOlePaieBSCalInfo
  else
    if RBPaieInfo.Checked then Result := GetVerbeOlePaieRendInfo
  else
    if RBEffectif.Checked then Result := GetVerbeOlePaieEffectif
  else
    if RBChoixMultisoc.Checked then   //PT1
    begin
      result := True;
      RBChoixMultisoc.Checked := False;
      RBBilansocial.Checked := True;
    end;
end;

procedure TFAssistExcelOle.PBC_BSPRESENTATIONChange(Sender: TObject);
begin
  inherited;
PBC_INDICATEURBS.Plus := 'AND PIL_BSPRESENTATION="'+PBC_BSPRESENTATION.Value+'" '
end;

procedure TFAssistExcelOle.TYPEVALEURChange(Sender: TObject);
begin
  inherited;
 RUBRIQUE.DataType := '';
 if TypeValeur.Value = 'MC' then
     RUBRIQUE.DataType := 'PGCUM'
 else
   if (TypeValeur.Value = 'BR') or (TypeValeur.Value = 'MR') then
     RUBRIQUE.DataType := 'PGREMUN'
 else
   if (TypeValeur.Value = 'BC') or (TypeValeur.Value = 'MS') or (TypeValeur.Value = 'MP') then
     RUBRIQUE.DataType := 'PGCOTIS';

 if TabValCal <> RUBRIQUE.DataType then RUBRIQUE.Value := '';
 TabValCal := RUBRIQUE.DataType;
 RUBRIQUE.enabled := (RUBRIQUE.DataType<>'');
end;

function TFAssistExcelOle.GetVerbeOleBilanSocial: boolean;
Var
  DateDeb,Datefin,Pres,Ind,Etab,Categ : String;
  Emploi,Coeff,Qualif,CodeStat        : String;
  TravailN1,TravailN2,TravailN3,TravailN4 : String;
  TabLibre1,TabLibre2,TabLibre3,TabLibre4 : String;
begin
  Result       := False;
  if not isValidDate(PBC_DATEDEBUT.Text) then begin error := 1; exit; end;
  if not isValidDate(PBC_DATEFIN.Text) then   begin error := 2; exit; end;
  if PBC_BSPRESENTATION.Value='' then         begin error := 3; exit; end;
  if PBC_INDICATEURBS.Value='' then           begin error := 4; exit; end;

  DateDeb      := PBC_DATEDEBUT.Text;
  Datefin      := PBC_DATEFIN.Text;
  Pres         := PBC_BSPRESENTATION.Value;
  Ind          := PBC_INDICATEURBS.Value;
  if Pos('<<',PBC_ETABLISSEMENT.Text) > 0 then Etab := ''      else Etab       := PBC_ETABLISSEMENT.Text;
  if Pos('<<',PBC_CATBILAN.Text)      > 0 then Categ := ''     else Categ      := PBC_CATBILAN.Text;
  if Pos('<<',PBC_LIBELLEEMPLOI.Text) > 0 then Emploi := ''    else Emploi     := PBC_LIBELLEEMPLOI.Text;
  if Pos('<<',PBC_COEFFICIENT.Text)   > 0 then Coeff := ''     else Coeff      := PBC_COEFFICIENT.Text;
  if Pos('<<',PBC_QUALIFICATION.Text) > 0 then Qualif := ''    else Qualif     := PBC_QUALIFICATION.Text;
  CodeStat := '';
  TravailN1:= '';  TravailN2:= '';  TravailN3:= '';  TravailN4:= '';
  TabLibre1:= '';  TabLibre2:= '';  TabLibre3:= '';  TabLibre4:= '';

//  if Pos('<<',PBC_CODESTAT.Value)      > 0 then CodeStat := ''  else CodeStat   := PBC_CODESTAT.Value;
//  if Pos('<<',PBC_TRAVAILN1.Value)     > 0 then TravailN1 := '' else TravailN1  := PBC_TRAVAILN1.Value;
//  if Pos('<<',PBC_TRAVAILN2.Value)     > 0 then TravailN2 := '' else TravailN2  := PBC_TRAVAILN2.Value;
//  if Pos('<<',PBC_TRAVAILN3.Value)     > 0 then TravailN3 := '' else TravailN3  := PBC_TRAVAILN3.Value;
//  if Pos('<<',PBC_TRAVAILN4.Value)     > 0 then TravailN4 := '' else TravailN4  := PBC_TRAVAILN4.Value;
//  if Pos('<<',PBC_LIBREPCMB1.Value)    > 0 then TabLibre1 := '' else TabLibre1  := PBC_LIBREPCMB1.Value;
//  if Pos('<<',PBC_LIBREPCMB2.Value)    > 0 then TabLibre2 := '' else TabLibre2  := PBC_LIBREPCMB2.Value;
//  if Pos('<<',PBC_LIBREPCMB3.Value)    > 0 then TabLibre3 := '' else TabLibre3  := PBC_LIBREPCMB3.Value;
//  if Pos('<<',PBC_LIBREPCMB4.Value)    > 0 then TabLibre4 := '' else TabLibre4  := PBC_LIBREPCMB4.Value;

  FormuleOle := '=PAIEBSINDOLE("'+DateDeb+'";"'+Datefin+'";"'+Pres+'";"'+Ind+'";"'+Etab+'";"'+
                               Categ+'";"'+Emploi+'";"'+Coeff+'";"'+Qualif+'";"'+CodeStat+'";"'+
                            TravailN1+'";"'+TravailN2+'";"'+TravailN3+'";"'+TravailN4+'";"'+
                            TabLibre1+'";"'+TabLibre2+'";"'+TabLibre3+'";"'+TabLibre4+'")';

  Result := True;
end;

function TFAssistExcelOle.GetVerbeOlePaieCumul: boolean;
Var
  Etab,Sal,TDonnee,TValeur,Rub,CatDucs,StatProf : String;
  YY,MM,DD : Word;
  ZmoisD, ZanneeD, ZmoisF, ZAnneeF, CodeStat,Pres : String;
  TravailN1,TravailN2,TravailN3,TravailN4 : String;
  TabLibre1,TabLibre2,TabLibre3,TabLibre4 : String;
begin
  Result       := False;
  if not isValidDate(DateDeb.Text) then begin error := 1; exit; end;
  if not isValidDate(DateFin.Text) then begin error := 2; exit; end;
  if TypeDonnee.Value=''           then begin error := 7; exit; end;
  if TypeValeur.Value=''           then begin error := 8; exit; end;
  if (TypeValeur.Value = 'MC') AND (RUBRIQUE.Value='') then begin error := 9; exit; end;
  if ((TypeValeur.Value = 'BR') or (TypeValeur.Value = 'MR'))
  AND (RUBRIQUE.Value='') then           begin error := 10; exit; end;
  if ((TypeValeur.Value = 'BC') or (TypeValeur.Value = 'MS') or (TypeValeur.Value = 'MP'))
  AND (RUBRIQUE.Value='') then           begin error := 11; exit; end;



  DeCodeDate(StrToDate(DateDeb.text), YY, MM, DD);
  ZanneeD := IntToStr(YY);
  ZmoisD  := IntToStr(MM);
  DeCodeDate(StrToDate(DateFin.text), YY, MM, DD);
  ZanneeF := IntToStr(YY);
  ZmoisF  := IntToStr(MM);
  Sal          := Salarie.Text;
  TDonnee      := TypeDonnee.Value;
  TValeur      := TypeValeur.Value;
  if Pos('<<',ETABLISSEMENT.Text) > 0 Then Etab:= ''       else Etab       := ETABLISSEMENT.Text;
  if Pos('<<',RUBRIQUE.Text)      > 0 Then Rub:= ''        else Rub        := RUBRIQUE.Text;
  if Pos('<<',CATDADS.Text)       > 0 Then CatDucs:= ''    else CatDucs    := CATDADS.Text;
  if Pos('<<',DADSPROF.Text)      > 0 Then StatProf:= ''   else StatProf   := DADSPROF.Text;
  if Pos('<<',MCODESTAT.Text)     > 0 then CodeStat  := '' else CodeStat   := MCODESTAT.Text;
  if Pos('<<',MTRAVAILN1.Text)    > 0 then TravailN1 := '' else TravailN1  := MTRAVAILN1.Text;
  if Pos('<<',MTRAVAILN2.Text)    > 0 then TravailN2 := '' else TravailN2  := MTRAVAILN2.Text;
  if Pos('<<',MTRAVAILN3.Text)    > 0 then TravailN3 := '' else TravailN3  := MTRAVAILN3.Text;
  if Pos('<<',MTRAVAILN4.Text)    > 0 then TravailN4 := '' else TravailN4  := MTRAVAILN4.Text;
  if Pos('<<',MLIBREPCMB1.Text)   > 0 then TabLibre1 := '' else TabLibre1  := MLIBREPCMB1.Text;
  if Pos('<<',MLIBREPCMB2.Text)   > 0 then TabLibre2 := '' else TabLibre2  := MLIBREPCMB2.Text;
  if Pos('<<',MLIBREPCMB3.Text)   > 0 then TabLibre3 := '' else TabLibre3  := MLIBREPCMB3.Text;
  if Pos('<<',MLIBREPCMB4.Text)   > 0 then TabLibre4 := '' else TabLibre4  := MLIBREPCMB4.Text;

  SupprimePointVirguleFin(Etab);
  SupprimePointVirguleFin(Rub);
  SupprimePointVirguleFin(CatDucs);
  SupprimePointVirguleFin(StatProf);
  SupprimePointVirguleFin(CodeStat);
  SupprimePointVirguleFin(TravailN1);   SupprimePointVirguleFin(TravailN2);
  SupprimePointVirguleFin(TravailN3);   SupprimePointVirguleFin(TravailN4);
  SupprimePointVirguleFin(TabLibre1);   SupprimePointVirguleFin(TabLibre2);
  SupprimePointVirguleFin(TabLibre3);   SupprimePointVirguleFin(TabLibre4);

  if Present.Checked = True then  Pres := 'O' else Pres := 'N';

  FormuleOle := '=PAIECUMUL("'+Etab+'";"'+Sal+'";"'+TDonnee+'";"'+TValeur+'";"'+Rub+'";"'+
                            TravailN1+'";"'+TravailN2+'";"'+TravailN3+'";"'+TravailN4+'";"'+
                            TabLibre1+'";"'+TabLibre2+'";"'+TabLibre3+'";"'+TabLibre4+'";"'+
                            CodeStat+'";"'+ZMoisD+'";"'+ZAnneeD+'";"'+ZMoisF+'";"'+
                            ZAnneeF+'";"'+CatDucs+'";"'+StatProf+'";"'+Pres+'")';
  Result := True;
end;

procedure TFAssistExcelOle.SALARIEExit(Sender: TObject);
begin
  inherited;
  if THCritMaskEdit(Sender).text <> '' then
    if (VH_PAIE.PgTypeNumSal='NUM') then Salarie.text := ColleZeroDevant(StrToInt(THCritMaskEdit(Sender).text),10);
end;


procedure TFAssistExcelOle.TYPERENSChange(Sender: TObject);
begin
  inherited;
  if not Assigned(TYPEVAL1) then exit;
  if not Assigned(VAL1) then exit;
  if not Assigned(VAL2) then exit;
  TYPEVAL1.Items.Clear;
  TYPEVAL1.Values.Clear;
  TYPEVAL1.DataType := '';
  TYPEVAL1.Value := '';
  VAL1.Enabled := False;
  VAL2.Enabled := False;
  VAL1.Text := '';
  VAL2.Text := '';
  if TypeRens.Value = 'EMP' then
    TYPEVAL1.DataType := 'PGCONDEMPLOI'
  else
     if TypeRens.Value = 'SEX' then
       TYPEVAL1.DataType := 'PGSEXE'
  else
     if TypeRens.Value = 'NAT' then
        Begin
        TYPEVAL1.Items.Add('Français');
        TYPEVAL1.Items.Add('Etranger');
        TYPEVAL1.Values.Add('FRA');
        TYPEVAL1.Values.Add('!FRA');
        End
  else
     if TypeRens.Value = 'ABS' then
        Begin
        TYPEVAL1.Items.Add('CP Pris');
        TYPEVAL1.Items.Add('Motif Absence 1 des paramètres sociétés');
        TYPEVAL1.Items.Add('Motif Absence 2 des paramètres sociétés');
        TYPEVAL1.Items.Add('Autres motifs');
        TYPEVAL1.Values.Add('0');
        TYPEVAL1.Values.Add('1');
        TYPEVAL1.Values.Add('2');
        TYPEVAL1.Values.Add('3');
        End
  else
     if (TypeRens.Value = 'AGE') OR (TypeRens.Value = 'ANC') then
        Begin
        TYPEVAL1.Enabled := FALSE;
        TYPEVAL1.Value := '';
        VAL1.Enabled := True;
        VAL2.Enabled := True;
        End
  else
     if (TypeRens.Value = 'CUM') then
          TYPEVAL1.DataType := 'PGCUMULPAIE';

end;



function TFAssistExcelOle.GetVerbeOlePaieBSCalInfo: boolean;
Var
  StVal1, StVal2 : String;
begin
  Result       := False;
  if not isValidDate(DateD.Text) then begin error := 1; exit; end;
  if not isValidDate(DateF.Text) then begin error := 2; exit; end;
  if TYPERENS.Value=''           then begin error := 7; exit; end;
  if (TYPERENS.Value='AGE') OR (TYPERENS.Value='ANC') then
    Begin
    IF Val1.text = '' then begin error := 12; exit; end;
    IF Val2.text = '' then begin error := 12; exit; end;
    StVal1 := Val1.Text;
    StVal2 := Val2.Text;
    End
  else
    Begin
    IF TypeVal1.Value = '' then begin error := 12; exit; end;
    StVal1 := TypeVal1.Value;
    End;
  FormuleOle := '=PAIEBSCALINFO("'+TYPERENS.Value+'";"'+StVal1+'";"'+StVal2+'";'+
                               '"'+DateD.Text+'";"'+DateF.Text+'";"'+CAT.Text+'")';
  Result := True;
end;

function TFAssistExcelOle.GetVerbeOlePaieRendInfo: boolean;
begin
  Result := False;
  if TYPETABLE.Value = '' then begin error := 14; exit; end;
  //if Salarie2.Text='' then begin error := 6; exit; end;
  if CHAMP.Value = '' then begin error := 13; exit; end;
  FormuleOle := '=PAIERENDINFO("'+TYPETABLE.Value+'";"'+SALARIE2.Text+
                               '";"'+Copy(CHAMP.Value,5,Length(CHAMP.Value))+'";'+
                               '"'+QUOI.Value+'")';
  Result := True;
end;

procedure TFAssistExcelOle.SupprimePointVirguleFin(var St: string);
Var i : integer;
begin
if St = '' then exit;
i := Length(St);
If St[i] = ';' then St := Copy(St,1,i-1);
end;

function TFAssistExcelOle.GetVerbeOlePaieEffectif: Boolean;
Var
  Etab,Sexe,CatDucs,DadsCat,StatProf,Contrat,Somme : String;
  YY,MM,DD : Word;
  ZmoisF, ZAnneeF, CodeStat,EntSort : String;
  LibEmploi,Coeff,Qualif,Indice,Niveau,CodeEmploi,MotifCont : string;
  TravailN1,TravailN2,TravailN3,TravailN4 : String;
  TabLibre1,TabLibre2,TabLibre3,TabLibre4 : String;
begin
  Result       := False;
  if not isValidDate(MENSUALITE.Text) then begin error := 2; exit; end;
  Somme := '';  EntSort := '';   MotifCont := '';

  DeCodeDate(StrToDate(MENSUALITE.text), YY, MM, DD);
  ZanneeF   := IntToStr(YY);
  ZmoisF    := IntToStr(MM);
  Sexe      := PSexe.Value;
  Contrat   := PTYPECONTRAT.Value;
  if Contrat <> '' then
    Begin
    EntSort   := PEntSort.Value;
    MotifCont := PMOTIFCONT.Value;
    If PSOMEFF.Value = 'OUI' then  Somme := '***';
    End;

  if Pos('<<',PETABLISSEMENT.Text) > 0 Then Etab:= ''        else Etab       := PETABLISSEMENT.Text;
  if Pos('<<',PCATDUCS.Text)       > 0 Then CatDucs:= ''     else CatDucs    := PCATDUCS.Text;
  if Pos('<<',PDADSPROF.Text)      > 0 Then StatProf:= ''    else StatProf   := PDADSPROF.Text;
  if Pos('<<',PDADSCAT.Text)       > 0 Then DadsCat:= ''     else DadsCat    := PDADSCAT.Text;

  if Pos('<<',PLIBELLEEMPLOI.Text) > 0 then LibEmploi := ''  else LibEmploi  := PLIBELLEEMPLOI.Text;
  if Pos('<<',PCOEFFICIENT.Text)   > 0 then Coeff := ''      else Coeff      := PCOEFFICIENT.Text;
  if Pos('<<',PINDICE.Text)        > 0 then Indice := ''     else Indice     := PINDICE.Text;
  if Pos('<<',PNIVEAU.Text)        > 0 then Niveau := ''     else Niveau     := PNIVEAU.Text;
  if Pos('<<',PQUALIFICATION.Text) > 0 then Qualif := ''     else Qualif     := PQUALIFICATION.Text;
  if Pos('<<',PCODEEMPLOI.Text)    > 0 then CodeEmploi := '' else CodeEmploi := PCODEEMPLOI.Text;

  if Pos('<<',MCODESTAT.Text)     > 0 then CodeStat  := '' else CodeStat   := MCODESTAT.Text;
  if Pos('<<',MTRAVAILN1.Text)    > 0 then TravailN1 := '' else TravailN1  := MTRAVAILN1.Text;
  if Pos('<<',MTRAVAILN2.Text)    > 0 then TravailN2 := '' else TravailN2  := MTRAVAILN2.Text;
  if Pos('<<',MTRAVAILN3.Text)    > 0 then TravailN3 := '' else TravailN3  := MTRAVAILN3.Text;
  if Pos('<<',MTRAVAILN4.Text)    > 0 then TravailN4 := '' else TravailN4  := MTRAVAILN4.Text;
  if Pos('<<',MLIBREPCMB1.Text)   > 0 then TabLibre1 := '' else TabLibre1  := MLIBREPCMB1.Text;
  if Pos('<<',MLIBREPCMB2.Text)   > 0 then TabLibre2 := '' else TabLibre2  := MLIBREPCMB2.Text;
  if Pos('<<',MLIBREPCMB3.Text)   > 0 then TabLibre3 := '' else TabLibre3  := MLIBREPCMB3.Text;
  if Pos('<<',MLIBREPCMB4.Text)   > 0 then TabLibre4 := '' else TabLibre4  := MLIBREPCMB4.Text;

  SupprimePointVirguleFin(Etab);       SupprimePointVirguleFin(CatDucs);
  SupprimePointVirguleFin(StatProf);   SupprimePointVirguleFin(DadsCat);
  SupprimePointVirguleFin(LibEmploi);
  SupprimePointVirguleFin(Coeff);      SupprimePointVirguleFin(Qualif);
  SupprimePointVirguleFin(Indice);     SupprimePointVirguleFin(CodeEmploi);
  SupprimePointVirguleFin(CodeStat);
  SupprimePointVirguleFin(TravailN1);  SupprimePointVirguleFin(TravailN2);
  SupprimePointVirguleFin(TravailN3);  SupprimePointVirguleFin(TravailN4);
  SupprimePointVirguleFin(TabLibre1);  SupprimePointVirguleFin(TabLibre2);
  SupprimePointVirguleFin(TabLibre3);  SupprimePointVirguleFin(TabLibre4);

  FormuleOle := '=PAIEEFFECTIF("'+Etab+'";"'+TravailN1+'";"'+TravailN2+'";"'+TravailN3+'";"'+TravailN4+'";"'+
                            TabLibre1+'";"'+TabLibre2+'";"'+TabLibre3+'";"'+TabLibre4+'";"'+
                            CodeStat+'";"'+ZMoisF+'";"'+ZAnneeF+'";"'+CatDucs+'";"'+StatProf+'";"'+
                            DadsCat+'";"'+Sexe+'";"'+LibEmploi+'";"'+Coeff+'";"'+Indice+'";"'+
                            Niveau+'";"'+Qualif+'";"'+CodeEmploi+'";"'+Contrat+'";"'+EntSort+'";"'+
                            MotifCont+'";"'+Somme+'")';
  Result := True;

end;



procedure TFAssistExcelOle.PENTSORTChange(Sender: TObject);
begin
  inherited;
if PEntSort.Value = 'E' then PMOTIFCONT.DataType :='PGMOTIFENTREELIGHT'
else if PEntSort.Value = 'S' then PMOTIFCONT.DataType :='PGMOTIFSORTIE';
PMOTIFCONT.Enabled := (PEntSort.Value<>'');
end;

procedure TFAssistExcelOle.PTYPECONTRATChange(Sender: TObject);
begin
  inherited;
if PTypeContrat.Value <> '' then
  Begin
  PEntSort.Enabled := True;
  PSOMEFF.Enabled := True;
  End
else
  Begin
  PEntSort.Enabled := False;
  PMOTIFCONT.Enabled := False;
  PSOMEFF.Enabled := False;
  PEntSort.Value:='';
  PMOTIFCONT.Value:='';
  PSOMEFF.Value:='';
  End;
end;
//PT1
procedure TFAssistExcelOle.MULTIDOSSIERSelect(Sender: TObject);
var
  ListeSoc, Soc : String;
begin
  inherited;
  ListeSoc := GetBasesMSWithLib(MULTIDOSSIER.Value);
  Soc := READTOKENPipe(ListeSoc,';');
  SELECTIONSOCIETE.Clear;
  while Soc <> '' do
  begin
    SELECTIONSOCIETE.Items.Add(READTOKENPipe(Soc,'|'));
    SELECTIONSOCIETE.Values.Add(Soc);
    Soc := READTOKENPipe(ListeSoc,';');
  end;

end;

procedure TFAssistExcelOle.RBBilansocialClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := False;
  TMULTIDOSSIER.visible := False;
  SELECTIONSOCIETE.visible := False;
  TSELECTIONSOCIETE.visible := False;
  Bsuivant.enabled := True;
  end;

procedure TFAssistExcelOle.RBChoixMultiSOcClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := True;
  TMULTIDOSSIER.visible := True;
  SELECTIONSOCIETE.visible := True;
  TSELECTIONSOCIETE.visible := True;
  Bsuivant.enabled := False;
end;

procedure TFAssistExcelOle.RBBilanSocialSimpleClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := False;
  TMULTIDOSSIER.visible := False;
  SELECTIONSOCIETE.visible := False;
  TSELECTIONSOCIETE.visible := False;
  Bsuivant.enabled := True;

end;

procedure TFAssistExcelOle.RBCumulPaieClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := False;
  TMULTIDOSSIER.visible := False;
  SELECTIONSOCIETE.visible := False;
  TSELECTIONSOCIETE.visible := False;
  Bsuivant.enabled := True;

end;

procedure TFAssistExcelOle.RBEffectifClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := False;
  TMULTIDOSSIER.visible := False;
  SELECTIONSOCIETE.visible := False;
  TSELECTIONSOCIETE.visible := False;
  Bsuivant.enabled := True;

end;

procedure TFAssistExcelOle.RBPaieInfoClick(Sender: TObject);
begin
  inherited;
  MULTIDOSSIER.visible := False;
  TMULTIDOSSIER.visible := False;
  SELECTIONSOCIETE.visible := False;
  TSELECTIONSOCIETE.visible := False;
  Bsuivant.enabled := True;

end;
//Fin PT1
end.


