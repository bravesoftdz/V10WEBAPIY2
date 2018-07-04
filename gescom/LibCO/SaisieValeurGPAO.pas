unit SaisieValeurGPAO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  UTOB, Hent1, ExtCtrls, Buttons, StdCtrls, Hctrls;

type
  TSaisieValeurDefaut = class(TForm)
    FBox: TScrollBox;
    Panels_On: TPanel;
    Panel1: TPanel;
    BAnnuler: TBitBtn;
    BOK: TBitBtn;
    procedure BAnnulerClick(Sender: TObject);
    procedure BOKClick(Sender: TObject);
  private
    { Déclarations privées }
    Procedure ChargeLaSaisie ;
    function  AffecteDataType(var NomChamp : string) : String;
    procedure Panels_OnEnter(Sender: TObject);
    function  LongueurSaisie (TypeChamp : string) : integer;
    procedure AjoutValeurCombo;
  public
    { Déclarations publiques }
    Tob_Valeur : TOB ;
  end;

var
  SaisieValeurDefaut: TSaisieValeurDefaut;


  Procedure AppelSaisieValeurGPAO (var Tob_Param : TOB);

implementation


////////////////////////////////////////////////////////////////////////////////
// Ouverture de la forme permettant de définir les valeurs par défaut des
// champs à reprendre
////////////////////////////////////////////////////////////////////////////////
Procedure AppelSaisieValeurGPAO (var Tob_Param : TOB);
var X  : TSaisieValeurDefaut ;
BEGIN
  SourisSablier;

  X:=TSaisieValeurDefaut.Create(Application) ;
  X.Tob_Valeur := Tob_Param ;
  X.ChargeLaSaisie;

  try
    X.ShowModal ;
  finally
    X.Free ;
  end ;
  SourisNormale ;
END ;


////////////////////////////////////////////////////////////////////////////////
// Balayage de la TOB des champs à reprendre
// Recherche du type de donnée, et formatage de la zone de saisie
////////////////////////////////////////////////////////////////////////////////
Procedure TSaisieValeurDefaut.ChargeLaSaisie ;
var Compteur : integer ;
    TC       : string ;
    TobFille : TOB     ;
    P  : TPanel;
    L  : TLabel;
    CC : TCheckBox;
    C  : THValComboBox;
    E  : THEdit;
    OkCase,OkCombo,OkEdit : boolean;
    i  : integer;
begin
  //
  // Il faut au moins un champ dans la TOB
  //
  if Tob_Valeur.detail.count <= 0 then exit ;

  for Compteur:=0 to Tob_Valeur.detail.count-1 do
  begin
    TobFille   := Tob_Valeur.Detail [Compteur];
    TC         := TobFille.GetValue ('GRC_TYPECHAMP');

    if TC='' then Exit ;
    OkCase:=(TC='BOOLEAN'); OkCombo:=(TC='COMBO');
    OkEdit:=((TC<>'BOOLEAN')And(TC<>'COMBO'));
    if ((Not OkCase)And(Not OkCombo)And(Not OkEdit)) then Exit;

    P:=TPanel.Create(Self); P.Parent:=FBox; P.Name:='P'+IntToStr(Compteur);
    P.ParentFont:=False; P.Font.Color:=clWindowText;
    P.Height:=30; P.Align:=alTop;
    P.BevelInner:=bvLowered; P.BevelOuter:=bvRaised;
    P.Caption:='';
    P.OnEnter:=Panels_OnEnter; P.OnClick:=Panels_OnEnter;
    P.Width:=Fbox.Width-40;
    //P.Width:=375;

    if OkCase then
    begin
      CC:=TCheckBox.Create(Self); CC.Parent:=P;
      //CC.Name:=TobFille.GetValue ('GRC_NOMCHAMP');
      CC.Name:='BOOL'+IntToStr(Compteur);
      CC.Left:=8; CC.Width:=300; CC.Alignment:= taLeftJustify; CC.Top:=4;
      CC.Caption:=TobFille.GetValue ('GRC_LIBELLECHAMP');
      if TobFille.GetValue ('GRC_VALDEFAUT') = 'X' then CC.Checked := True
      else CC.Checked := False;
      //CC.setfocus;
    end else
    begin
      L:=THLabel.Create(Self); L.Parent:=P;
      L.Name:='L'+IntToStr(Compteur);
      L.Left:=8; L.Width:=150; L.Top:=8; L.AutoSize:=false;
      L.Caption:=TobFille.GetValue ('GRC_LIBELLECHAMP');
      if OkCombo then
      begin
        P.Caption := P.Name;
        C:=THValComboBox.Create(Self);
        C.Parent:=P;
        //C.Name:=TobFille.GetValue ('GRC_NOMCHAMP');
        C.Name:='COM'+IntToStr(Compteur);
        C.Left:=L.Left+L.Width+1;
        C.Width:=Fbox.Width-40-C.Left-8;
        C.Top := 4 ;
        C.Text:='' ;
        C.Style:=csDropDownList ;
        //C.DataType:=AffecteDataType(NomChamp);
        //if TobFille.GetValue ('GRC_VALDEFAUT') <> '' then C.Value := TobFille.GetValue ('GRC_VALDEFAUT')
        //else begin
        //  C.ItemIndex:=-1;
        //  C.Value := '';
        //end;
      end;
      if OkEdit then
      begin
        E:=THEdit.Create(Self); E.Parent:=P;
        //E.Name:=TobFille.GetValue ('GRC_NOMCHAMP');
        E.Name:='EDIT'+IntToStr(Compteur);
        if TC='DATE' then
        begin
          E.EditMask:='##/##/####';
          E.OpeType := otDate;
        end else if ((TC='INTEGER') or (TC='DOUBLE')) then E.OpeType := otReel
        else begin
          E.OpeType   := otString;
          E.MaxLength := LongueurSaisie (TobFille.GetValue ('GRC_TYPECHAMP'));
        end;
        E.Left:=L.Left+L.Width+1;
        //E.Width:=P.Width-E.Left-8;
        //E.Width:=Fbox.Width-40-C.Left-8;
        E.Width:=172;
        E.Top:=4;
        E.Text:=''; //L.FocusControl:=E;
        if TobFille.GetValue ('GRC_VALDEFAUT') <> '' then E.text := TobFille.GetValue ('GRC_VALDEFAUT');
        //E.SetFocus;
      end;
    end;

    for i:=0 to FBox.ControlCount-1 do if FBox.Controls[i] is TPanel then
    begin
      if TPanel(FBox.Controls[i])=P then TPanel(FBox.Controls[i]).Color:=clTeal
      else TPanel(FBox.Controls[i]).Color:=clBtnFace;
    end;
  end;

  AjoutValeurCombo ;
end;

///////////////////////////////////////////////////////////////////////////////
// Ajoute les valeurs dans les combos
///////////////////////////////////////////////////////////////////////////////
procedure TSaisieValeurDefaut.AjoutValeurCombo;
var Tob_Fille : TOB     ;
    i,j       : integer ;
    NumTob    : integer ;
    P         : TPanel  ;
    St        : string  ;
    NomChamp  : string  ;
begin

  for i:=0 to FBox.ControlCount-1 do
  if ((FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On')) then
  begin
    P:=TPanel(FBox.Controls[i]);
    St:=P.Name;
    if Pos('P',St)<=0 then break;
    NumTob := StrToInt(Copy(St,2,Length(St)));

    for j:=0 to P.ControlCount-1 do
    begin
      if P.Controls[j] is THValComboBox then
      begin
        Tob_Fille := Tob_Valeur.detail[NumTob];
        NomChamp  := Tob_Fille.GetValue ('GRC_NOMCHAMP');
        if Nomchamp = 'GDI_TYPEDIM' then THValComboBox(P.Controls[j]).DataType:='GCCATEGORIEDIM'
        else THValComboBox(P.Controls[j]).DataType:=AffecteDataType(NomChamp);
        if Tob_Fille.GetValue ('GRC_VALDEFAUT') <> '' then THValComboBox(P.Controls[j]).Value := Tob_Fille.GetValue ('GRC_VALDEFAUT')
        else begin
          THValComboBox(P.Controls[j]).ItemIndex:=-1;
          THValComboBox(P.Controls[j]).Value := '';
        end;
      end;
      if (P.Controls[j] is THEdit) and (THEDIT(P.Controls[j]).OpeType = otString) then
      begin
        Tob_Fille := Tob_Valeur.detail[NumTob];
        NomChamp  := Tob_Fille.GetValue ('GRC_NOMCHAMP');
        if Nomchamp = 'CC_CODE' then St:= ''
        else St := AffecteDataType(NomChamp) ;
        if St <> '' then
        begin
          THEdit(P.Controls[j]).DataType:=St;
          THEdit(P.Controls[j]).ElipsisButton := True;
        end;
      end;
    end;
  end;
end;



function TSaisieValeurDefaut.AffecteDataType(var NomChamp : string) : String;
BEGIN
Result:='';
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Result:=Get_Join(NomChamp);
{$ENDIF}
END;

function TSaisieValeurDefaut.LongueurSaisie (TypeChamp : string) : integer;
var tempo : string ;
BEGIN
  Tempo := Copy (TypeChamp, 1, 7);
  if Tempo<>'VARCHAR' then result := 35        // Il y a un PB, on renvoie la valeur par défaut
  else begin
    Tempo  := Strpos (PChar(TypeChamp), '(');
    result := StrToInt (copy (Tempo, 2, length (Tempo)-2));
   end;
END;

procedure TSaisieValeurDefaut.Panels_OnEnter(Sender: TObject);
var St     : string;
    i,IP,j : integer;
    Curlig : integer;
    P      : TPanel;
begin
  St:=TComponent(Sender).Name; if Pos('P',St)<=0 then exit;

  CurLig:=StrToInt(Copy(St,2,Length(St)));
  for i:=0 to FBox.ControlCount-1 do
  if ((FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On')) then
  begin
    P:=TPanel(FBox.Controls[i]); St:=P.Name; if Pos('P',St)<=0 then break;
    IP:=StrToInt(Copy(St,2,Length(St)));
    if IP=CurLig then
    begin
      P.Color:=clTeal;
      for j:=0 to P.ControlCount do
      begin
        if P.Controls[j] is TCheckBox then BEGIN TCheckBox(P.Controls[j]).setfocus; break; END;
        if P.Controls[j] is THEdit then BEGIN THEdit(P.Controls[j]).setfocus; break; END;
        if P.Controls[j] is THValComboBox then BEGIN THValComboBox(P.Controls[j]).setfocus; break; END;
      end;
    end else P.Color:=clBtnFace;
  end;
end;

{$R *.DFM}
///////////////////////////////////////////////////////////////////////////////
// On sort sans sauvegarde
///////////////////////////////////////////////////////////////////////////////
procedure TSaisieValeurDefaut.BAnnulerClick(Sender: TObject);
begin
Self.Close;
end;

///////////////////////////////////////////////////////////////////////////////
// Sauvegarde des données
///////////////////////////////////////////////////////////////////////////////
procedure TSaisieValeurDefaut.BOKClick(Sender: TObject);
var
    i,j     : integer ;
    NumTob  : integer ;
    P       : TPanel;
    St : string;
begin

  for i:=0 to FBox.ControlCount-1 do
  if ((FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On')) then
  begin
    P:=TPanel(FBox.Controls[i]);
    St:=P.Name;
    if Pos('P',St)<=0 then break;
    NumTob := StrToInt(Copy(St,2,Length(St)));

    for j:=0 to P.ControlCount-1 do
    begin
      if P.Controls[j] is TCheckBox then
      begin
        if TCheckBox(P.Controls[j]).checked then Tob_Valeur.detail[NumTob].PutValue ('GRC_VALDEFAUT', 'X')
        else Tob_Valeur.detail[NumTob].PutValue ('GRC_VALDEFAUT', '-');
      end
      else if P.Controls[j] is THEdit then
      begin
       Tob_Valeur.detail[NumTob].PutValue ('GRC_VALDEFAUT', THEdit(P.Controls[j]).Text);
      end
      else if P.Controls[j] is THValComboBox then
      begin
        Tob_Valeur.detail[NumTob].PutValue ('GRC_VALDEFAUT', THValComboBox(P.Controls[j]).value);
      end;
    end;
  end;
end;

end.
