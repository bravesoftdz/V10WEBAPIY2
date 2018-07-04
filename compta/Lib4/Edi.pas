unit EDI;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, StdCtrls, TabNotBk, Buttons, ExtCtrls,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB,
  Mask, Hctrls, Ent1, ComCtrls, HEnt1, hmsgbox, Hcompte, HSysMenu, Hqry, ParamSoc,
  ADODB;

type PlanCpt = (cptGEN,cptAUX,cptANA,cptBUD) ;

type TDossier = RECORD
  OkSauve : Boolean ;
  EditRefDem : String ;
  CombFctTxt : String3 ;
  EditRefTxt,EditTxt,EditNumDosEmt,EditNumDosRec,EditIdentEnv:String ;
  // Echange
  EditEmet,EditAdrEmet,EditRec,EditAdrRec,EditRefIntChg : String ;
  CBoxAccRec,CBoxEssai:String1 ;
  // Dossier
  DEditSiret,DEditFctCont,DEditPerChrg,DEditNumCom,DEditModCom,DEditNom,DEditRue1,
  DEditRue2,DEditVille,DEditDiv,DEditCodPost,DEditPays : String ;
  //Emetteur
  EEditSiret,EEditFctCont,EEditPerChrg,EEditNumCom,EEditModCom,EEditNom,EEditRue1,
  EEditRue2,EEditVille,EEditDiv,EEditCodPost,EEditPays : String ;
  //Recepteur
  RComboPerChrg : String ;
  REditSiret,REditFctCont,REditNumCom,REditModCom,REditNom,REditRue1,
  REditRue2,REditVille,REditDiv,REditCodPost,REditPays : String ;
  end ;

var Dossier : TDossier ;

type
  TFEDI = class(TForm)
    Panel1       : TPanel;
    BtAide: THBitBtn;
    BtSortir: THBitBtn;
    BtValider: THBitBtn;
    MsgBox: THMsgBox;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    GroupBox1: TGroupBox;
    LabRefDem: THLabel;
    LabFctTxt: THLabel;
    LabRefTxt: THLabel;
    LabTxt: THLabel;
    LabNumDosEmt: THLabel;
    LabNumDosRec: THLabel;
    LabIdentEnv: THLabel;
    EditRefDem: TEdit;
    CombFctTxt: TComboBox;
    EditRefTxt: TEdit;
    EditTxt: TEdit;
    EditNumDosEmt: TEdit;
    EditNumDosRec: TEdit;
    EditIdentEnv: TEdit;
    Label5: TLabel;
    Groupdos: TGroupBox;
    DLabSiret: TLabel;
    DLabFctCont: TLabel;
    DLabPerChrg: TLabel;
    DLabRue1: TLabel;
    DLabDiv: TLabel;
    DLabCodPost: TLabel;
    DLabPays: TLabel;
    DLabVille: TLabel;
    DLabRue2: TLabel;
    DLabNumCom: TLabel;
    DLabModCom: TLabel;
    DLabNom: TLabel;
    Label2: TLabel;
    DEditSiret: TEdit;
    DEditFctCont: TEdit;
    DEditPerChrg: TEdit;
    DEditRue1: TEdit;
    DEditDiv: TEdit;
    DEditCodPost: TEdit;
    DEditPays: TEdit;
    DEditVille: TEdit;
    DEditRue2: TEdit;
    DEditNumCom: TEdit;
    DEditModCom: TEdit;
    DEditNom: TEdit;
    GroupEmet: TGroupBox;
    ELabSiret: TLabel;
    ELabFctCont: TLabel;
    ELabPerChrg: TLabel;
    ELabNumCom: TLabel;
    ELabModCom: TLabel;
    ELabRue1: TLabel;
    ELabRue2: TLabel;
    ELabVille: TLabel;
    ELabDiv: TLabel;
    ELabCodPost: TLabel;
    ELabPays: TLabel;
    ELabNom: TLabel;
    Label3: TLabel;
    EEditSiret: TEdit;
    EEditFctCont: TEdit;
    EEditPerChrg: TEdit;
    EEditNumCom: TEdit;
    EEditModCom: TEdit;
    EEditRue1: TEdit;
    EEditRue2: TEdit;
    EEditVille: TEdit;
    EEditDiv: TEdit;
    EEditCodPost: TEdit;
    EEditPays: TEdit;
    EEditNom: TEdit;
    GroupRec: TGroupBox;
    RLabSiret: TLabel;
    RLabFctCont: TLabel;
    RLabPerChrg: TLabel;
    RLabNumCom: TLabel;
    RLabModCom: TLabel;
    RLabRue1: TLabel;
    RLabRue2: TLabel;
    RLabVille: TLabel;
    RLabDiv: TLabel;
    RLabCodPost: TLabel;
    RLabPays: TLabel;
    RLabNom: TLabel;
    Label4: TLabel;
    REditSiret: TEdit;
    REditFctCont: TEdit;
    REditNumCom: TEdit;
    REditModCom: TEdit;
    REditRue1: TEdit;
    REditRue2: TEdit;
    REditVille: TEdit;
    REditDiv: TEdit;
    REditCodPost: TEdit;
    REditPays: TEdit;
    REditNom: TEdit;
    RComboPerChrg: TComboBox;
    RComboTel: TComboBox;
    Query: TQuery;
    EditEmet: TEdit;
    LabEmet: TLabel;
    LabAdrRec: TLabel;
    EditAdrEmet: TEdit;
    EditRec: TEdit;
    EditRec1: THCpteEdit;
    EditAdrRec: TEdit;
    LabAdrEmet: TLabel;
    LabRec: TLabel;
    EEditRec: TLabel;
    LabRefIntChg: TLabel;
    EditRefIntChg: TEdit;
    CBoxEssai: TCheckBox;
    CBoxAccRec: TCheckBox;
    TFMTCHOIX: THTable;
    TFMTCHOIXIC_IMPORT: TStringField;
    TFMTCHOIXIC_NATURE: TStringField;
    TFMTCHOIXIC_FORMAT: TStringField;
    TFMTCHOIXIC_DETRUIRE: TStringField;
    TFMTCHOIXIC_RESUME: TStringField;
    TFMTCHOIXIC_VALIDE: TStringField;
    TFMTCHOIXIC_RECYCLE: TStringField;
    TFMTCHOIXIC_FICHIER: TStringField;
    TFMTCHOIXIC_ASCII: TStringField;
    TFMTCHOIXIC_FIXEFORMAT: TStringField;
    TFMTCHOIXIC_FIXEFICHIER: TStringField;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BtSortirClick(Sender: TObject);
    procedure EditRec1Exit(Sender: TObject);
    procedure BtValiderClick(Sender: TObject);
    procedure RComboPerChrgChange(Sender: TObject);
    procedure BtAideClick(Sender: TObject);
  private
    Imp : String ;
    CurNat : String ;
    procedure SauveDossier ;
    procedure Charge ;
    procedure Sauve ;
    procedure NewEnreg ;
  public
  end;

procedure ParametresEDI(Imp : String ; CurNat : String) ;
function MTF(mtt: double;Precision,Chiffre: integer) : string ;
function NBS(nbr,taille: integer;car: char) : string ;
function Blanc(size: integer) : string ;
function QS(ch: string;lg: integer) : string ;
function QI(ch: Integer;chfiller: string) : string ;
function QF(ch: Double;Precision,Chiffre: integer) : string ;
function QD(ch: TDateTime;chformat: string) : string ;
function QB(ch: boolean) : boolean ;
Procedure InitDossier(Imp,CurNat : String) ;

implementation

{$R *.DFM}

procedure ParametresEDI(Imp : String ; CurNat : String) ;
var FEDI:TFEDI ;
BEGIN
FEDI:=TFEDI.Create(Application) ;
try
  FEDI.Imp:=Imp ;
  FEDI.CurNat:=CurNat ;
  FEDI.ShowModal ;
  finally
  FEDI.Free ;
  END ;
Screen.Cursor:=SynCrDefault ;
END;

procedure TFEDI.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Charge ;
end;

procedure TFEDI.BtSortirClick(Sender: TObject);
begin Close ; end;

procedure TFEDI.EditRec1Exit(Sender: TObject);
Var MonCompte : String ;
    i     : integer ;
begin
EditRec1.ExisteH ;
EditRec.text:=EEditRec.Caption ;
if EditRec1.Text='' then exit ;
Query.Close ;
Query.sql.Clear ;
Query.sql.add('Select T_AUXILIAIRE,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,'+
              'T_CODEPOSTAL,T_VILLE,T_DIVTERRIT,T_PAYS,T_TELEPHONE,'+
              'T_SIRET from TIERS where T_AUXILIAIRE = "'+EditRec.Text+'"') ;
ChangeSQL(Query) ; //Query.Prepare ;
PrepareSQLODBC(Query) ;
Query.Open ;
REditSiret.Text:=QS(Query.findField('T_SIRET').AsString,0) ;
REditNom.Text:=QS(Query.findField('T_LIBELLE').asString,0) ;
REditRue1.Text:=QS(Query.findField('T_ADRESSE1').asString,0) ;
REditRue2.Text:=QS(Query.findField('T_ADRESSE2').asString,0) ;
REditVille.Text:=QS(Query.findField('T_VILLE').asString,0) ;
REditDiv.Text:=QS(Query.findField('T_DIVTERRIT').asString,0) ;
REditCodPost.Text:=QS(Query.findField('T_CODEPOSTAL').asString,0) ;
REditPays.Text:=QS(Query.findField('T_PAYS').asString,2) ;
MonCompte:=QS(Query.findField('T_AUXILIAIRE').asString,0) ;
Query.Close ;
Query.sql.Clear ;
//Simon "
Query.sql.add('Select C_NOM,C_PRINCIPAL,C_SERVICE,C_TELEPHONE,C_FONCTION '+
              'from CONTACT where C_AUXILIAIRE = "'+MonCompte+'"') ;
ChangeSQL(Query) ; //Query.Prepare ;
PrepareSQLODBC(Query) ;
Query.Open ;
Query.First ;
RComboPerChrg.Items.Clear ;
RComboTel.Items.Clear ;
REditNumCom.Text:='' ;
i:=0;
while not Query.EOF do
  BEGIN
  RComboPerChrg.Items.Add(QS(Query.FindField('C_NOM').AsString,0)) ;
  RComboTel.Items.Add(QS(Query.FindField('C_TELEPHONE').AsString,0)) ;
  if QS(Query.FindField('C_PRINCIPAL').AsString,1)='O' then
     BEGIN
     RComboPerChrg.ItemIndex:=i ;
     REditNumCom.Text:=QS(Query.FindField('C_TELEPHONE').AsString,0) ;
     END ;
  Inc(i) ;
  Query.Next ;
  END ;
Query.Close;
end;

procedure TFEDI.RComboPerChrgChange(Sender: TObject);
begin
REditNumCom.Text:=RComboTel.Items[RComboPerChrg.ItemIndex] ;
end;

procedure TFEDI.BtValiderClick(Sender: TObject);
BEGIN
SauveDossier ;
Close ;
end;

{------------------------ OK --------------------------}
function Blanc(size: integer) : string ;
var i   : integer ;
    str : string ;
BEGIN
str:='' ;
for i:=1 to size do
  str:=str+' ' ;
Blanc:=str ;
END ;

{------------------------ OK --------------------------}
function QS(ch: string;lg: integer) : string ;
BEGIN
if lg=0 then QS:=ch
        else QS:=Format_String(copy(ch,1,lg),lg) ;
END ;

{------------------------ OK --------------------------}
function QI(ch: Integer;chfiller: string) : string ;
BEGIN
QI:=FormatFloat(chfiller,ch) ;
END ;

{------------------------ OK --------------------------}
function QF(ch: Double;Precision,Chiffre: integer) : string ;
var res : string ;
    i   : integer ;
BEGIN
res:='' ;
for i:=1 to Precision+1-Length(FloatToStrF(ch,ffFixed,Precision,chiffre)) do res:=res+' ' ;
QF:=res+FloatToStrF(ch,ffFixed,Precision,chiffre) ;
END ;

{------------------------ OK --------------------------}
function QD(ch: TDateTime;chformat: string) : string ;
var res : string ;
BEGIN
res:=FormatDateTime(chformat,ch) ;
if res='00000000' then res:='        ' ;
QD:=Res ;
END ;

{------------------------ OK --------------------------}
function QB(ch: boolean) : boolean ;
BEGIN
QB:=ch ;
END ;

{------------------------ OK --------------------------}
function MTF(mtt: double;Precision,Chiffre: integer) : string ;
var res : string ;
    i   : integer ;
BEGIN
res:='' ;
for i:=1 to Precision+1-Length(FloatToStrF(mtt,ffFixed,Precision,chiffre)) do
  res:=res+' ' ;
MTF:=res+FloatToStrF(mtt,ffFixed,Precision,chiffre) ;
END ;

{------------------------ OK --------------------------}
function NBS(nbr,Taille: integer;car: char) : string ;
var res : string ;
    i   : integer ;
BEGIN
res:='' ;
for i:=1 to taille-Length(IntToStr(nbr)) do
  res:=res+car ;
NBS:=Res+IntToStr(nbr) ;
END ;

Procedure InitDossier(Imp,CurNat : String) ;
var Q1 : TQuery ;
    St,Nam,Val : String ;
    Lines : TStringList ;
    j  : integer ;
BEGIN
Q1:=OpenSQL('Select IC_EDI From FMTCHOIX Where IC_IMPORT="'+Imp+'" AND IC_NATURE="'+CurNat+'"',True) ;
if Q1.EOF then BEGIN Ferme(Q1) ; Exit ; END ; 
if (Q1.Fields[0].AsString<>'') then
  BEGIN
  Lines:=TStringList.Create ;
  Lines.Assign(TMemoField(Q1.FindField('IC_EDI'))) ;
  With Dossier do
    BEGIN
    OkSauve:=True ;
    for j:=0 to Lines.Count-1 do
      BEGIN
      St:=Lines[j] ; Nam:=ReadTokenSt(St) ; Val:=ReadTokenSt(St) ;
      if Nam='EditRefDem' then EditRefDem:=Val else
      if Nam='CombFctTxt' then CombFctTxt:=Val else
      if Nam='EditRefTxt' then EditRefTxt:=Val else
      if Nam='EditTxt' then EditTxt:=Val else
      if Nam='EditNumDosEmt' then EditNumDosEmt:=Val else
      if Nam='EditNumDosRec' then EditNumDosRec:=Val else
      if Nam='EditIdentEnv' then EditIdentEnv:=Val else
      // Echange
      if Nam='EditEmet' then EditEmet:=Val else
      if Nam='EditAdrEmet' then EditAdrEmet:=Val else
      if Nam='EditRec' then EditRec:=Val else
      if Nam='EditAdrRec' then EditAdrRec:=Val else
      if Nam='EditRefIntChg' then EditRefIntChg:=Val else
      if Nam='CBoxAccRec' then CBoxAccRec:=Val else
      if Nam='CBoxEssai' then CBoxEssai:=Val else
      // Dossier
      if Nam='DEditSiret' then DEditSiret:=Val else
      if Nam='DEditFctCont' then DEditFctCont:=Val else
      if Nam='DEditPerChrg' then DEditPerChrg:=Val else
      if Nam='DEditNumCom' then DEditNumCom:=Val else
      if Nam='DEditModCom' then DEditModCom:=Val else
      if Nam='DEditNom' then DEditNom:=Val else
      if Nam='DEditRue1' then DEditRue1:=Val else
      if Nam='DEditRue2' then DEditRue2:=Val else
      if Nam='DEditVille' then DEditVille:=Val else
      if Nam='DEditDiv' then DEditDiv:=Val else
      if Nam='DEditCodPost' then DEditCodPost:=Val else
      if Nam='DEditPays' then DEditPays:=Val else
      //Emetteur
      if Nam='EEditSiret' then EEditSiret:=Val else
      if Nam='EEditFctCont' then EEditFctCont:=Val else
      if Nam='EEditPerChrg' then EEditPerChrg:=Val else
      if Nam='EEditNumCom' then EEditNumCom:=Val else
      if Nam='EEditModCom' then EEditModCom:=Val else
      if Nam='EEditNom' then EEditNom:=Val else
      if Nam='EEditRue1' then EEditRue1:=Val else
      if Nam='EEditRue2' then EEditRue2:=Val else
      if Nam='EEditVille' then EEditVille:=Val else
      if Nam='EEditDiv' then EEditDiv:=Val else
      if Nam='EEditCodPost' then EEditCodPost:=Val else
      if Nam='EEditPays' then EEditPays:=Val else
      //Recepteur
      if Nam='REditSiret' then REditSiret:=Val else
      if Nam='REditFctCont' then REditFctCont:=Val else
      if Nam='RComboPerChrg' then RComboPerChrg:=Val else
      if Nam='REditNumCom' then REditNumCom:=Val else
      if Nam='REditModCom' then REditModCom:=Val else
      if Nam='REditNom' then REditNom:=Val else
      if Nam='REditRue1' then REditRue1:=Val else
      if Nam='REditRue2' then REditRue2:=Val else
      if Nam='REditVille' then REditVille:=Val else
      if Nam='REditDiv' then REditDiv:=Val else
      if Nam='REditCodPost' then REditCodPost:=Val else
      if Nam='REditPays' then REditPays:=Val ;
      END ;
    Lines.Free ;
    END ;
  Ferme(Q1) ;
  Exit ;
  END Else BEGIN Ferme(Q1) ; Exit ; END ;
{$IFDEF SPEC302}
Q:=OpenSQL('Select SO_SIRET,SO_CONTACT,SO_LIBELLE,SO_ADRESSE1,'+
              'SO_ADRESSE2,SO_CODEPOSTAL,SO_VILLE,SO_DIVTERRIT,SO_PAYS,'+
              'SO_TELEPHONE From SOCIETE',True) ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
With Dossier do
  BEGIN
  EditEmet:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
  EEditNom:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
  EEditSiret:=QS(Q.FindField('SO_SIRET').AsString,0) ;
  EEditPerChrg:=QS(Q.FindField('SO_CONTACT').AsString,0) ;
  EEditNumCom:=QS(Q.FindField('SO_TELEPHONE').AsString,0) ;
  EEditRue1:=QS(Q.FindField('SO_ADRESSE1').AsString,0) ;
  EEditRue2:=QS(Q.FindField('SO_ADRESSE2').AsString,0) ;
  EEditVille:=QS(Q.FindField('SO_VILLE').AsString,0) ;
  EEditDiv:=QS(Q.FindField('SO_DIVTERRIT').AsString,0) ;
  EEditCodPost:=QS(Q.FindField('SO_CODEPOSTAL').AsString,0) ;
  EEditPays:=Copy(QS(Q.FindField('SO_PAYS').AsString,0),1,2) ;
  EEditFctCont:='AU ' ;
  EEditModCom:='TE ' ;
  DEditNom:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
  DEditSiret:=QS(Q.FindField('SO_SIRET').AsString,0) ;
  DEditPerChrg:=QS(Q.FindField('SO_CONTACT').AsString,0) ;
  DEditNumCom:=QS(Q.FindField('SO_TELEPHONE').AsString,0) ;
  DEditRue1:=QS(Q.FindField('SO_ADRESSE1').AsString,0) ;
  DEditRue2:=QS(Q.FindField('SO_ADRESSE2').AsString,0) ;
  DEditVille:=QS(Q.FindField('SO_VILLE').AsString,0) ;
  DEditDiv:=QS(Q.FindField('SO_DIVTERRIT').AsString,0) ;
  DEditCodPost:=QS(Q.FindField('SO_CODEPOSTAL').AsString,0) ;
  DEditPays:=QS(Q.FindField('SO_PAYS').AsString,2) ;
  DEditFctCont:='AU ' ;
  DEditModCom:='TE ' ;
  REditFctCont:='AU ' ;
  REditModCom:='TE ' ;
  END ;
Ferme(Q) ;
{$ELSE}
With Dossier do
  BEGIN
  EditEmet:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
  EEditNom:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
  EEditSiret:=QS(GetParamSocSecur('SO_SIRET',''),0) ;
  EEditPerChrg:=QS(GetParamSocSecur('SO_CONTACT',''),0) ;
  EEditNumCom:=QS(GetParamSocSecur('SO_TELEPHONE',''),0) ;
  EEditRue1:=QS(GetParamSocSecur('SO_ADRESSE1',''),0) ;
  EEditRue2:=QS(GetParamSocSecur('SO_ADRESSE2',''),0) ;
  EEditVille:=QS(GetParamSocSecur('SO_VILLE',''),0) ;
  EEditDiv:=QS(GetParamSocSecur('SO_DIVTERRIT',''),0) ;
  EEditCodPost:=QS(GetParamSocSecur('SO_CODEPOSTAL',''),0) ;
  EEditPays:=Copy(QS(GetParamSocSecur('SO_PAYS',''),0),1,2) ;
  EEditFctCont:='AU ' ;
  EEditModCom:='TE ' ;
  DEditNom:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
  DEditSiret:=QS(GetParamSocSecur('SO_SIRET',''),0) ;
  DEditPerChrg:=QS(GetParamSocSecur('SO_CONTACT',''),0) ;
  DEditNumCom:=QS(GetParamSocSecur('SO_TELEPHONE',''),0) ;
  DEditRue1:=QS(GetParamSocSecur('SO_ADRESSE1',''),0) ;
  DEditRue2:=QS(GetParamSocSecur('SO_ADRESSE2',''),0) ;
  DEditVille:=QS(GetParamSocSecur('SO_VILLE',''),0) ;
  DEditDiv:=QS(GetParamSocSecur('SO_DIVTERRIT',''),0) ;
  DEditCodPost:=QS(GetParamSocSecur('SO_CODEPOSTAL',''),0) ;
  DEditPays:=QS(GetParamSocSecur('SO_PAYS',''),2) ;
  DEditFctCont:='AU ' ;
  DEditModCom:='TE ' ;
  REditFctCont:='AU ' ;
  REditModCom:='TE ' ;
  END ;
{$ENDIF}
END ;

procedure TFEDI.SauveDossier;
BEGIN
With Dossier do
  BEGIN
  OkSauve:=True ;
  EditRefDem :=TEdit(FindComponent('EditRefDem')).Text ;
  CombFctTxt :=TComboBox(FindComponent('CombFctTxt')).Text;
  EditRefTxt :=TEdit(FindComponent('EditRefTxt')).Text ;
  EditTxt:=TEdit(FindComponent('EditTxt')).Text ;
  EditNumDosEmt:=TEdit(FindComponent('EditNumDosEmt')).Text ;
  EditNumDosRec:=TEdit(FindComponent('EditNumDosRec')).Text ;
  EditIdentEnv:=TEdit(FindComponent('EditIdentEnv')).Text ;
  // Echange
  EditEmet:=TEdit(FindComponent('EditEmet')).Text ;
  EditAdrEmet:=TEdit(FindComponent('EditAdrEmet')).Text ;
  EditRec:=TEdit(FindComponent('EditRec')).Text ;
  EditAdrRec:=TEdit(FindComponent('EditAdrRec')).Text ;
  EditRefIntChg:=TEdit(FindComponent('EditRefIntChg')).Text ;
  if (TCheckBox(FindComponent('CBoxAccRec')).checked) then CBoxAccRec:='1' else CBoxAccRec:=' ' ;
  if (TCheckBox(FindComponent('CBoxEssai')).Checked )then CBoxEssai:='1' else CBoxEssai:=' ' ;
  // Dossier
  DEditSiret:=TEdit(FindComponent('DEditSiret')).Text ;
  DEditFctCont:=TEdit(FindComponent('DEditFctCont')).Text ;
  DEditPerChrg:=TEdit(FindComponent('DEditPerChrg')).Text ;
  DEditNumCom:=TEdit(FindComponent('DEditNumCom')).Text ;
  DEditModCom:=TEdit(FindComponent('DEditModCom')).Text ;
  DEditNom:=TEdit(FindComponent('DEditNom')).Text ;
  DEditRue1:=TEdit(FindComponent('DEditRue1')).Text ;
  DEditRue2:=TEdit(FindComponent('DEditRue2')).Text ;
  DEditVille:=TEdit(FindComponent('DEditVille')).Text ;
  DEditDiv:=TEdit(FindComponent('DEditDiv')).Text ;
  DEditCodPost:=TEdit(FindComponent('DEditCodPost')).Text ;
  DEditPays:=TEdit(FindComponent('DEditPays')).Text ;
  //Emetteur
  EEditSiret:=TEdit(FindComponent('EEditSiret')).Text ;
  EEditFctCont:=TEdit(FindComponent('EEditFctCont')).Text ;
  EEditPerChrg:=TEdit(FindComponent('EEditPerChrg')).Text ;
  EEditNumCom:=TEdit(FindComponent('EEditNumCom')).Text ;
  EEditModCom:=TEdit(FindComponent('EEditModCom')).Text ;
  EEditNom:=TEdit(FindComponent('EEditNom')).Text ;
  EEditRue1:=TEdit(FindComponent('EEditRue1')).Text ;
  EEditRue2:=TEdit(FindComponent('EEditRue2')).Text ;
  EEditVille:=TEdit(FindComponent('EEditVille')).Text ;
  EEditDiv:=TEdit(FindComponent('EEditDiv')).Text ;
  EEditCodPost:=TEdit(FindComponent('EEditCodPost')).Text ;
  EEditPays:=TEdit(FindComponent('EEditPays')).Text ;
  //Recepteur
  REditSiret:=TEdit(FindComponent('REditSiret')).Text ;
  REditFctCont:=TEdit(FindComponent('REditFctCont')).Text ;
  RComboPerChrg:=TComboBox(FindComponent('RComboPerChrg')).Text ;
  REditNumCom:=TEdit(FindComponent('REditNumCom')).Text ;
  REditModCom:=TEdit(FindComponent('REditModCom')).Text ;
  REditNom:=TEdit(FindComponent('REditNom')).Text ;
  REditRue1:=TEdit(FindComponent('REditRue1')).Text ;
  REditRue2:=TEdit(FindComponent('REditRue2')).Text ;
  REditVille:=TEdit(FindComponent('REditVille')).Text ;
  REditDiv:=TEdit(FindComponent('REditDiv')).Text ;
  REditCodPost:=TEdit(FindComponent('REditCodPost')).Text ;
  REditPays:=TEdit(FindComponent('REditPays')).Text ;
  END ;
if Transactions(Sauve,5)<>OeOk then MessageAlerte(MsgBox.Mess[0]) ;
END ;

Procedure TFEDI.Charge ;
Var j : integer ;
     St,Nam : string ;
     C : TComponent ;
     Lines : TStringList ;

   procedure TextToControl ( St : String ; C : TControl ) ;
   BEGIN
   if C is THCritMaskEdit then THCritMaskEdit(C).Text:=St else
   if C is TEdit then TEdit(C).Text:=St else
   if C is TMaskEdit then TMaskEdit(C).Text:=St else
   if C is THValComboBox then
      BEGIN
      if ((THVaLComboBox(C).Vide) and (St='')) then THVaLComboBox(C).ItemIndex:=0
                                               else THVaLComboBox(C).Value:=St ;
      END else
   if C is THValComboBox then TComboBox(C).Text:=St else
   if C is TComboBox then TComboBox(C).Text:=St else
   if C is TCheckBox then TCheckBox(C).State:=TCheckBoxState(StrToInt(st)) ;
   END ;

BEGIN
Query.Close ;
Query.sql.Clear ;
Query.sql.add('Select * From FMTCHOIX Where IC_IMPORT="'+Imp+'" AND IC_NATURE="'+CurNat+'"') ;
ChangeSQL(Query) ; //Query.Prepare ;
PrepareSQLODBC(Query) ;
Query.Open ;
if Query.EOF then if Transactions(NewEnreg,5)<>OeOk then BEGIN MessageAlerte(MsgBox.Mess[0]) ; Exit ; END ;
Lines:=TStringList.Create ; Lines.Assign(TMemoField(Query.FindField('IC_EDI'))) ;
for j:=0 to Lines.Count-1 do
  BEGIN
  St:=Lines[j] ; Nam:=ReadTokenSt(St) ;
  C:=FindComponent(Nam) ;
  if C<>Nil then TextToControl(St,TControl(C)) ;
  END ;
Lines.Free ;
END ;

procedure TFEDI.NewEnreg ;
{$IFDEF SPEC302}
var Q : TQuery ;
{$ENDIF}
BEGIN
Query.Insert ;
Query.FindField('IC_IMPORT').AsString:=Imp ;
Query.FindField('IC_NATURE').AsString:=CurNat ;
Query.Post ;
InitDossier(Imp,CurNat) ;
{$IFDEF SPEC302}
Q:=OpenSQL('Select SO_SIRET,SO_CONTACT,SO_LIBELLE,SO_ADRESSE1,'+
              'SO_ADRESSE2,SO_CODEPOSTAL,SO_VILLE,SO_DIVTERRIT,SO_PAYS,'+
              'SO_TELEPHONE From SOCIETE',True) ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
EditEmet.Text:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
EEditNom.Text:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
EEditSiret.Text:=QS(Q.FindField('SO_SIRET').AsString,0) ;
EEditPerChrg.Text:=QS(Q.FindField('SO_CONTACT').AsString,0) ;
EEditNumCom.Text:=QS(Q.FindField('SO_TELEPHONE').AsString,0) ;
EEditRue1.Text:=QS(Q.FindField('SO_ADRESSE1').AsString,0) ;
EEditRue2.Text:=QS(Q.FindField('SO_ADRESSE2').AsString,0) ;
EEditVille.Text:=QS(Q.FindField('SO_VILLE').AsString,0) ;
EEditDiv.Text:=QS(Q.FindField('SO_DIVTERRIT').AsString,0) ;
EEditCodPost.Text:=QS(Q.FindField('SO_CODEPOSTAL').AsString,0) ;
EEditPays.Text:=Copy(QS(Q.FindField('SO_PAYS').AsString,0),1,2) ;
EEditFctCont.Text:='AU ' ;
EEditModCom.Text:='TE ' ;
DEditNom.Text:=QS(Q.FindField('SO_LIBELLE').AsString,0) ;
DEditSiret.Text:=QS(Q.FindField('SO_SIRET').AsString,0) ;
DEditPerChrg.Text:=QS(Q.FindField('SO_CONTACT').AsString,0) ;
DEditNumCom.Text:=QS(Q.FindField('SO_TELEPHONE').AsString,0) ;
DEditRue1.Text:=QS(Q.FindField('SO_ADRESSE1').AsString,0) ;
DEditRue2.Text:=QS(Q.FindField('SO_ADRESSE2').AsString,0) ;
DEditVille.Text:=QS(Q.FindField('SO_VILLE').AsString,0) ;
DEditDiv.Text:=QS(Q.FindField('SO_DIVTERRIT').AsString,0) ;
DEditCodPost.Text:=QS(Q.FindField('SO_CODEPOSTAL').AsString,0) ;
DEditPays.Text:=QS(Q.FindField('SO_PAYS').AsString,2) ;
DEditFctCont.Text:='AU ' ;
DEditModCom.Text:='TE ' ;
REditFctCont.Text:='AU ' ;
REditModCom.Text:='TE ' ;
Ferme(Q) ;
{$ELSE}
EditEmet.Text:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
EEditNom.Text:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
EEditSiret.Text:=QS(GetParamSocSecur('SO_SIRET',''),0) ;
EEditPerChrg.Text:=QS(GetParamSocSecur('SO_CONTACT',''),0) ;
EEditNumCom.Text:=QS(GetParamSocSecur('SO_TELEPHONE',''),0) ;
EEditRue1.Text:=QS(GetParamSocSecur('SO_ADRESSE1',''),0) ;
EEditRue2.Text:=QS(GetParamSocSecur('SO_ADRESSE2',''),0) ;
EEditVille.Text:=QS(GetParamSocSecur('SO_VILLE',''),0) ;
EEditDiv.Text:=QS(GetParamSocSecur('SO_DIVTERRIT',''),0) ;
EEditCodPost.Text:=QS(GetParamSocSecur('SO_CODEPOSTAL',''),0) ;
EEditPays.Text:=Copy(QS(GetParamSocSecur('SO_PAYS',''),0),1,2) ;
EEditFctCont.Text:='AU ' ;
EEditModCom.Text:='TE ' ;
DEditNom.Text:=QS(GetParamSocSecur('SO_LIBELLE',''),0) ;
DEditSiret.Text:=QS(GetParamSocSecur('SO_SIRET',''),0) ;
DEditPerChrg.Text:=QS(GetParamSocSecur('SO_CONTACT',''),0) ;
DEditNumCom.Text:=QS(GetParamSocSecur('SO_TELEPHONE',''),0) ;
DEditRue1.Text:=QS(GetParamSocSecur('SO_ADRESSE1',''),0) ;
DEditRue2.Text:=QS(GetParamSocSecur('SO_ADRESSE2',''),0) ;
DEditVille.Text:=QS(GetParamSocSecur('SO_VILLE',''),0) ;
DEditDiv.Text:=QS(GetParamSocSecur('SO_DIVTERRIT',''),0) ;
DEditCodPost.Text:=QS(GetParamSocSecur('SO_CODEPOSTAL',''),0) ;
DEditPays.Text:=QS(GetParamSocSecur('SO_PAYS',''),2) ;
DEditFctCont.Text:='AU ' ;
DEditModCom.Text:='TE ' ;
REditFctCont.Text:='AU ' ;
REditModCom.Text:='TE ' ;
{$ENDIF}
END ;

Procedure TFEDI.Sauve ;
Var j,i,k : integer ;
    St  : string ;
    C,CP   : TControl ;
    Lines : TStringList ;

   Procedure ControlToText ( C : TControl ; Var St : String ) ;
   BEGIN
   if C is THCritMaskEdit then St:=THCritMaskEdit(C).Text else
   if C is TEdit then St:=TEdit(C).Text else
   if C is TMaskEdit then St:=TMaskEdit(C).Text else
   if C is THValComboBox then St:=THvalComboBox(C).Value else
   if C is THValComboBox then St:=TComboBox(C).Text else
   if C is TComboBox then St:=TComboBox(C).Text else
   if C is TCheckBox then St:=InttoStr(Integer(TCheckBox(C).State)) ;
   END ;

BEGIN
Query.Close ;
Query.sql.Clear ;
Query.sql.add('Select * From FMTCHOIX Where IC_IMPORT="'+Imp+'" AND IC_NATURE="'+CurNat+'"') ;
ChangeSQL(Query) ; //Query.Prepare ;
PrepareSQLODBC(Query) ;
Query.Open ;
if Query.EOF then Exit ;
Query.Edit ;
Lines:=TStringList.Create ;
Lines.Clear ;
for j:=0 to Pages.PageCount-1 do for i:=0 to Pages.Pages[j].ControlCount-1 do
  BEGIN
  C:=Pages.Pages[j].Controls[i] ;
  if C is TGroupBox then
    BEGIN
    if (C.Visible) and (C.Enabled) then
      for k:=0 to TGroupBox(C).ControlCount-1 do
        BEGIN
        CP:=TGroupBox(C).Controls[k] ;
        St:='@' ; ControlToText(CP,St) ;
        if st<>'@' then Lines.Add(CP.Name+';'+St) ;
        END ;
    END else
    BEGIN
    St:='@' ; ControlToText(C,St) ;
    if st<>'@' then Lines.Add(C.Name+';'+St) ;
    END ;
  END ;
TMemoField(Query.FindField('IC_EDI')).Assign(Lines) ;
Query.Post ;
Lines.Free ;
END ;

procedure TFEDI.BtAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

