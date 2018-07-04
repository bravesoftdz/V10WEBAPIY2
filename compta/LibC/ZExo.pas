unit ZExo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Hctrls, HSysMenu, hmsgbox, Mask,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  UTob, HEnt1,                                    { +- Bibli  -+ }
  Ent1,                                           { +- Compta -+ }
  TZ ;                                            { +-  PFU   -+ }

function SaisieZExo(NbExo : Integer; DateFin : TDateTime) : Boolean ;

type
  TZFExo = class(TForm)
    Panel1: TPanel;
    panel: TPanel;
    HelpBtn: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    TEX_EXERCICE: THLabel;
    TEX_ABREGE: THLabel;
    TEX_LIBELLE: THLabel;
    TEX_DATEDEBUT: THLabel;
    TEX_DATEFIN: THLabel;
    EX_EXERCICE: TEdit;
    EX_ABREGE: TEdit;
    EX_LIBELLE: TEdit;
    EX_DATEFIN: TEdit;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    EX_DATEDEBUT: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure EX_LIBELLEExit(Sender: TObject);
  private
    NbExo   : Integer ;
    DateFin : TDateTime ;
    function ControleSaisie(bVerbose : Boolean) : Boolean ;
  public
    { Déclarations publiques }
  end;

type
  TZExo = class
  private
  Exos : TZF ;
  function    GetCount : Integer ;
  public
  constructor Create ;
  destructor  Destroy ; override ;
  function    Load : Boolean ;
  function    Find(Exo : string) : Integer ;
  function    GetFirst : string ;
  function    GetPrev(Exo : string) : Integer ;
  function    GetNext(Exo : string) : Integer ;
  function    GetLast : string ;
  function    GetValue(Nom : string; Niveau : integer) : Variant ;
  property    Count : Integer read GetCount ;
  end ;

implementation

{$R *.DFM}

//=======================================================
//======== Point d'entrée dans la saisie devise =========
//=======================================================
function SaisieZExo(NbExo : Integer; DateFin : TDateTime) : Boolean ;
var ZFExo : TZFExo ;
begin
ZFExo:=TZFExo.Create(Application) ;
  try
  ZFExo.DateFin:=DateFin ; ZFExo.NbExo:=NbExo ;
  Result:=(ZFExo.ShowModal=mrOK) ;
  finally
  ZFExo.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
function TZFExo.ControleSaisie(bVerbose : Boolean) : Boolean ;
begin
Result:=TRUE ;
if EX_ABREGE.Text='' then
   begin if bVerbose then HM.Execute(0,'','') ; Result:=FALSE ; Exit ; end ;
if EX_LIBELLE.Text='' then
   begin if bVerbose then HM.Execute(1,'','') ; Result:=FALSE ; Exit ; end ;
if (EX_DATEDEBUT.Text='') or
   (StrToDate(EX_DATEDEBUT.Text)>=StrToDate(EX_DATEFIN.Text)) then
   begin if bVerbose then HM.Execute(2,'','') ; Result:=FALSE ; Exit ; end ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TZFExo.FormShow(Sender: TObject) ;
var DateDeb : TDateTime ; i : Integer ;
begin
EX_EXERCICE.Text:=Format('%.3d', [NbExo+1]) ;
EX_DATEFIN.Text:=DateToStr(DateFin) ;
DateDeb:=DateFin ;
for i:=1 to 12 do
    begin DateDeb:=DebutDeMois(DateDeb) ; if i<>12 then DateDeb:=DateDeb-1 ; end ;
EX_DATEDEBUT.Text:=DateToSTr(DateDeb) ;
EX_LIBELLE.SetFocus ;
end ;

procedure TZFExo.EX_LIBELLEExit(Sender: TObject);
begin
if EX_ABREGE.Text='' then
 begin
 if Length(EX_LIBELLE.Text)>17 then EX_ABREGE.Text:=Trim(Copy(EX_LIBELLE.Text, 1, 17))
                               else EX_ABREGE.Text:=EX_LIBELLE.Text ;
 end ;
end;

procedure TZFExo.BValiderClick(Sender: TObject) ;
var Q : TQuery ;
begin
if not ControleSaisie(TRUE) then Exit ;
Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+W_W+'"', FALSE) ;
Q.Insert ; InitNew(Q) ;
Q.FindField('EX_EXERCICE').AsString:=EX_EXERCICE.Text ;
Q.FindField('EX_LIBELLE').AsString:=EX_LIBELLE.Text ;
Q.FindField('EX_ABREGE').AsString:=EX_ABREGE.Text ;
Q.FindField('EX_DATEDEBUT').AsDateTime:=StrToDate(EX_DATEDEBUT.Text) ;
Q.FindField('EX_DATEFIN').AsDateTime:=StrToDate(EX_DATEFIN.Text) ;
Q.FindField('EX_ETATCPTA').AsString:='CDE' ;
Q.FindField('EX_ETATBUDGET').AsString:='OUV' ;
Q.FindField('EX_SOCIETE').AsString:=V_PGI.CodeSociete ;
Q.FindField('EX_VALIDEE').AsString:='-------------------------' ;
Q.Post ;
Ferme(Q) ;
CHARGEMAGEXO(FALSE) ;
ModalResult:=mrOk ;
end;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// EXERCICE : EX_EXERCICE

constructor TZExo.Create ;
begin
Exos:=TZF.Create('VEXERCICE', nil, -1) ;
end ;

destructor TZExo.Destroy ;
begin
if Exos<>nil then Exos.Free ;
end ;

function TZExo.Load : Boolean ;
var Q : TQuery ; Col : string ;
begin
Col:='EX_EXERCICE, EX_ABREGE, EX_LIBELLE, EX_DATEDEBUT, EX_DATEFIN' ;
Q:=OpenSQL('SELECT '+Col+' FROM EXERCICE ORDER BY EX_DATEDEBUT', TRUE);
Result:=not Q.EOF ; if not Result then begin Ferme(Q) ; Exit ; end ;
Result:=Exos.LoadDetailDB('VEXOS', '', '', Q, FALSE) ;
Ferme(Q) ;
end ;

function TZExo.GetCount : Integer ;
begin
Result:=Exos.Detail.Count ;
end ;

function TZExo.Find(Exo : string) : Integer ;
var i : integer ;
begin
Result:=-1 ;
Exo:=Trim(Exo) ; if Exo='' then Exit ;
for i:=0 to Exos.Detail.Count-1 do
  if TZF(Exos.Detail[i]).GetValue('EX_EXERCICE')=Exo then
     begin Result:=i ; Exit ; end ;
end ;

function TZExo.GetFirst : string ;
begin
Result:='' ;
if Exos.Detail.Count>0 then Result:=TZF(Exos.Detail[0]).GetValue('EX_EXERCICE') ;
end ;

function TZExo.GetLast : string ;
begin
Result:='' ;
if Exos.Detail.Count>0 then Result:=TZF(Exos.Detail[Exos.Detail.Count-1]).GetValue('EX_EXERCICE') ;
end ;

function TZExo.GetPrev(Exo : string) : Integer ;
var k : Integer ;
begin
Result:=-1 ; k:=Find(Exo) ;
if k>0 then Result:=k-1 ;
end ;

function TZExo.GetNext(Exo : string) : Integer ;
var k : Integer ;
begin
Result:=-1 ; k:=Find(Exo) ;
if (k>=0) and (k+1<=Exos.Detail.Count) then Result:=k+1 ;
end ;

function TZExo.GetValue(Nom : string; Niveau : integer) : Variant ;
begin
Result:=TZF(Exos.Detail[Niveau]).GetValue(Nom) ;
end ;

end.
