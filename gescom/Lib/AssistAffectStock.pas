unit AssistAffectStock;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HPanel, ComCtrls, HSysMenu, hmsgbox, StdCtrls, HTB97, ExtCtrls,
  Hctrls,UTOB, Mask,HEnt1, UtilGC;

  procedure Assist_AffectStock (TOBT, TOBS : TOB; Format : string);

type
  TFAssistStock = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    PBevel1: TBevel;
    TINTRO: THLabel;
    GBPRIOR: TGroupBox;
    PRIORITE: THValComboBox;
    TPRIORITE: THLabel;
    CBASCENDANT: TCheckBox;
    DATELIVRMAX: THCritMaskEdit;
    TDATELIVRMAX: TLabel;
    CBRELIQUATLIG: TCheckBox;
    TabSheet2: TTabSheet;
    PBevel2: TBevel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    HStk: THMsgBox;
    HMsgErr: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure DATELIVRMAXExit(Sender: TObject);
  private
    { Déclarations privées }
    TOBAff,TOBStk : TOB ;
    Format : string ;
//    function  SpaceStr ( nb : integer) : string;
//    function  ExtractLibelle ( St : string) : string;
    procedure ListeRecap;
    procedure AffecteStockAuto;
  public
    { Déclarations publiques }
  end;

var
  FAssistStock: TFAssistStock;
  i_NumEcran : integer;

implementation

{$R *.DFM}

procedure Assist_AffectStock (TOBT, TOBS : TOB; Format : string);
var
   Fo_Assist : TFAssistStock;
Begin
     Fo_Assist := TFAssistStock.Create (Application);
     Fo_Assist.TOBAff:= TOBT;
     Fo_Assist.TOBStk:= TOBS;
     Fo_Assist.Format:= Format;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

{=========================================================================================}
{============================= Evenements de la forme ====================================}
{=========================================================================================}
procedure TFAssistStock.FormShow(Sender: TObject);
var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
bAnnuler.Visible := True;
bSuivant.Enabled := True;
bFin.Visible := True;
bFin.Enabled := False;
//i_NumEcran := 0;

Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
PRIORITE.ItemIndex:=0;
end;

procedure TFAssistStock.DATELIVRMAXExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DATELIVRMAX.Text) then
  begin
    stErr:='"'+ DATELIVRMAX.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DATELIVRMAX.SetFocus;
  end;
end;

{=========================================================================================}
{================================= Récapitulatif =========================================}
{=========================================================================================}
(* function TFAssistStock.SpaceStr ( nb : integer) : string;
Var St_Chaine : string ;
    i_ind : integer ;
BEGIN
St_Chaine := '' ;
for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
Result:=St_Chaine;
END;

function TFAssistStock.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
BEGIN
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    BEGIN
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    END else St_Chaine := St;
Result := St_Chaine + ' : ';
END; *)

procedure TFAssistStock.ListeRecap;
Var st_chaine : string;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (GBPRIOR.Caption));
if CBASCENDANT.Checked then st_chaine := HStk.Mess[2] else  st_chaine := HStk.Mess[3];;
ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TPRIORITE.Caption) + PRIORITE.Text +
                     SpaceStr(4)  + st_chaine);
ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TDATELIVRMAX.Caption) + DATELIVRMAX.Text);
if CBRELIQUATLIG.Checked then st_chaine := HStk.Mess[0] else  st_chaine := HStk.Mess[1];;
ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBRELIQUATLIG.Caption) + st_chaine);
END;

{=========================================================================================}
{========================== Evenements Liés aux Boutons ==================================}
{=========================================================================================}

procedure TFAssistStock.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
if bFin.Enabled then ListeRecap;
end;

procedure TFAssistStock.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFAssistStock.bFinClick(Sender: TObject);
begin
  inherited;
AffecteStockAuto;
Close ;
end;

{=========================================================================================}
{========================== Evenements Liés aux Boutons ==================================}
{=========================================================================================}
procedure TFAssistStock.AffecteStockAuto;
var i_ind, i_inddeb, i_indfin, i_inc : integer;
    TOBL, TOBD : TOB ;
    RefArt,Depot,St : string;
    QteDispo, QteStk, QteAnc, QteNec : double;
    DateLiv : TDatetime;
begin
if TOBAff.Detail.count=0 then exit;
TOBAff.Detail.Sort(PRIORITE.Value) ;
if CBASCENDANT.Checked then
    BEGIN
    i_inddeb:=0; i_indfin:=TOBAff.Detail.count-1; i_inc:=1;
    END else
    BEGIN
    i_indfin:=0; i_inddeb:=TOBAff.Detail.count-1; i_inc:=-1;
    END;
i_ind:=i_inddeb-i_inc;
Repeat
i_ind:=i_ind+i_inc;
TOBL:=TOBAff.Detail[i_ind];
DateLiv:=TOBL.GetValue ('GL_DATELIVRAISON');
if DateLiv > StrToDate(DATELIVRMAX.Text) then Continue;
if TOBL.GetValue('GL_TENUESTOCK')<>'X' then
    BEGIN
    TOBL.PutValue('QTEAFF', TOBL.GetValue('GL_QTESTOCK'));
    Continue;
    END;
QteStk:=TOBL.GetValue ('GL_QTESTOCK'); if QteStk < 0 then Continue;
RefArt:=TOBL.GetValue('GL_ARTICLE');
Depot:=TOBL.GetValue('GL_DEPOT');
TOBD:=TOBStk.FindFirst(['ARTICLE','DEPOT'],[RefArt,Depot],False);
if TOBD=Nil then Continue;
QteAnc:=TOBL.GetValue ('QTEAFF');
QteNec:= QteStk - QteAnc;
QteDispo:=TOBD.GetValue('QTEDISPO'); if QteDispo < 0 then Continue;
if QteDispo < QteNec then
    if CBRELIQUATLIG.Checked then QteNec:=QteDispo else QteNec:=0;
TOBD.PutValue('QTEDISPO', QteDispo-QteNec);
St:=FormatFloat(Format, QteNec+QteAnc);
TOBL.PutValue ('QTEAFF', St);
until i_ind=i_indfin;
end;

end.
