unit SaisieValDefautASCII;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFNDEF EAGLCLIENT}
     dbtables,
  {$ENDIF}
  UTOB, Hent1, ExtCtrls, Buttons, StdCtrls, Hctrls;

type
  TFSaisieValDefautASCII = class(TForm)
    FBox: TScrollBox;
    Panels_On: TPanel;
    Panel1: TPanel;
    BAnnuler: TBitBtn;
    BOK: TBitBtn;
    procedure BAnnulerClick(Sender: TObject);
    procedure BOKClick(Sender: TObject);
  private
    { Déclarations privées }
    Tob_Champs : TOB;
    procedure ChargeLesChamps;
    function  RechercheChampSupp(TobFille : TOB; NomChamp : String): string;
    procedure Panels_OnEnter(Sender: TObject);
    function  LongueurSaisie (TypeChamp : string) : integer;
    procedure AjoutValeurCombo;
  public
    { Déclarations publiques }
  end;

var
  FSaisieValDefautASCII: TFSaisieValDefautASCII;

procedure AppelSaisieValDefautASCII(var Tob_Param : TOB);

implementation

{$R *.DFM}

////////////////////////////////////////////////////////////////////////////////
// Ouvre une fenêtre avec une liste de champs pour lesquels on pourra définir
// une valeur par défaut
////////////////////////////////////////////////////////////////////////////////
Procedure AppelSaisieValDefautASCII (var Tob_Param : TOB) ;
var X : TFSaisieValDefautASCII ;
begin
  SourisSablier;
  X := TFSaisieValDefautASCII.Create(Application) ;
  X.Tob_Champs := Tob_Param ;
  X.ChargeLesChamps ;
  try
    X.ShowModal;
  finally
    X.Free;
  end ;
  SourisNormale;
end ;

////////////////////////////////////////////////////////////////////////////////
// Balayage de la TOB des champs à reprendre
// Recherche du type de donnée, et formatage de la zone de saisie
////////////////////////////////////////////////////////////////////////////////
Procedure TFSaisieValDefautASCII.ChargeLesChamps ;
var Compteur,i : integer ;
    TC       : string ;
    TobFille : TOB ;
    P  : TPanel ;
    L  : TLabel ;
    CC : TCheckBox ;
    C  : THValComboBox ;
    E  : THEdit ;
    OkCase, OkCombo, OkEdit : boolean ;
begin
  if Tob_Champs.detail.count <= 0 then exit ;

  for Compteur := 0 to Tob_Champs.detail.count - 1 do
    begin
    TobFille := Tob_Champs.Detail[Compteur] ;

    TC := RechercheChampSupp (TobFille, 'TYPECHAMP') ; if TC = '' then Exit ;
    OkCase := (TC = 'BOOLEAN') ; OkCombo := (TC = 'COMBO') ;
    OkEdit := ((TC <> 'BOOLEAN') And (TC <> 'COMBO')) ;
    if ((Not OkCase) And (Not OkCombo) And (Not OkEdit)) then Exit ;

    P:=TPanel.Create(Self); P.Parent:=FBox; P.Name:='P'+IntToStr(Compteur);
    P.ParentFont:=False; P.Font.Color:=clWindowText;
    P.Height:=30; P.Align:=alTop;
    P.BevelInner:=bvLowered; P.BevelOuter:=bvRaised;
    P.Caption:='';
    P.OnEnter:=Panels_OnEnter; P.OnClick:=Panels_OnEnter;
    P.Width:=Fbox.Width-40;

    if OkCase then
      begin
      CC:=TCheckBox.Create(Self); CC.Parent:=P;
      CC.Name:='BOOL'+IntToStr(Compteur);
      CC.Left:=6; CC.Width:=323; CC.Alignment:=taLeftJustify; CC.Top:=7;
      CC.Caption:=RechercheChampSupp(TobFille,'LIBELLECHAMP');
      if TobFille.GetValue('GEX_VALDEFAUT')='X' then CC.Checked:=True
      else CC.Checked:=False;
      end
    else
      begin
      L:=THLabel.Create(Self); L.Parent:=P;
      L.Name:='L'+IntToStr(Compteur);
      L.Left:=8; L.Width:=150; L.Top:=8; L.AutoSize:=false;
      L.Caption:=RechercheChampSupp(TobFille,'LIBELLECHAMP');
      if OkCombo then
        begin
        P.Caption := P.Name;
        C:=THValComboBox.Create(Self);
        C.Parent:=P;
        C.Name:='COM'+IntToStr(Compteur);
        C.Left:=L.Left+L.Width+1;
        C.Width:=172;
        C.Top := 4;
        C.Text:= '';
        C.Style:=csDropDownList;
        C.Vide:=True ;
        C.VideString := TraduireMemoire ('<<Aucun>>') ;
        end;
      if OkEdit then
        begin
        E:=THEdit.Create(Self); E.Parent:=P;
        E.Name:='EDIT'+IntToStr(Compteur);
        if TC='DATE' then
          begin
          E.EditMask:='##/##/####';
          E.OpeType := otDate;
          end
        else if ((TC='INTEGER') or (TC='DOUBLE')) then E.OpeType := otReel
        else
          begin
          E.OpeType   := otString;
          E.MaxLength := LongueurSaisie (RechercheChampSupp(TobFille,'TYPECHAMP'));
          end;
        E.Left:=L.Left+L.Width+1;
        E.Width:=172;
        E.Top:=4;
        E.Text:='';
        if TobFille.GetValue('GEX_VALDEFAUT')<>'' then E.text:=TobFille.GetValue ('GEX_VALDEFAUT');
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

function TFSaisieValDefautASCII.RechercheChampSupp(TobFille : TOB; NomChamp : String): string;
var Q : TQuery;
    First : Boolean;
begin
First := Not TobFille.FieldExists(NomChamp);
if Not TobFille.FieldExists('TYPECHAMP') then TobFille.AddChampSup('TYPECHAMP',False);
if Not TobFille.FieldExists('LIBELLECHAMP') then TobFille.AddChampSup('LIBELLECHAMP',False);
if First then
  begin
  Q := OpenSQL('SELECT DH_TYPECHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_NOMCHAMP="'+TobFille.GetValue('GEX_NOMCHAMP')+'"',True) ;
  TobFille.PutValue('TYPECHAMP',Q.FindField('DH_TYPECHAMP').AsString);
  TobFille.PutValue('LIBELLECHAMP',Q.FindField('DH_LIBELLE').AsString);
  Ferme(Q);
  end;
Result := TobFille.GetValue(NomChamp);
end;

///////////////////////////////////////////////////////////////////////////////
// Ajoute les valeurs dans les combos
///////////////////////////////////////////////////////////////////////////////
procedure TFSaisieValDefautASCII.AjoutValeurCombo;
var Tob_Fille : TOB;
    i,j,NumTob : integer;
    P : TPanel;
    St,NomChamp : string;
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
        Tob_Fille := Tob_Champs.detail[NumTob];
        NomChamp  := Tob_Fille.GetValue('GEX_NOMCHAMP');
        THValComboBox(P.Controls[j]).DataType:=Get_Join(NomChamp);
        if Tob_Fille.GetValue('GEX_VALDEFAUT') <> '' then
          THValComboBox(P.Controls[j]).Value := Tob_Fille.GetValue ('GEX_VALDEFAUT')
        else begin
          THValComboBox(P.Controls[j]).ItemIndex:=-1;
          THValComboBox(P.Controls[j]).Value := '';
        end;
      end;
      if (P.Controls[j] is THEdit) and (THEDIT(P.Controls[j]).OpeType = otString) then
      begin
        Tob_Fille := Tob_Champs.detail[NumTob];
        NomChamp  := Tob_Fille.GetValue ('GEX_NOMCHAMP');
        St := Get_Join(NomChamp);
        if St <> '' then
        begin
          THEdit(P.Controls[j]).DataType:=St;
          THEdit(P.Controls[j]).ElipsisButton := True;
        end;
      end;
    end;
  end;
end;

function TFSaisieValDefautASCII.LongueurSaisie (TypeChamp : string) : integer;
var tempo : string ;
BEGIN
  Tempo := Copy (TypeChamp, 1, 7);
  if Tempo<>'VARCHAR' then result := 35        // Il y a un PB, on renvoie la valeur par défaut
  else begin
    Tempo  := Strpos (PChar(TypeChamp), '(');
    result := StrToInt (copy (Tempo, 2, length (Tempo)-2));
   end;
END;

procedure TFSaisieValDefautASCII.Panels_OnEnter(Sender: TObject);
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


procedure TFSaisieValDefautASCII.BAnnulerClick(Sender: TObject);
begin
Self.Close;
end;

procedure TFSaisieValDefautASCII.BOKClick(Sender: TObject);
var i,j,NumTob : integer;
    P : TPanel;
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
        if TCheckBox(P.Controls[j]).checked then Tob_Champs.detail[NumTob].PutValue('GEX_VALDEFAUT','X')
        else Tob_Champs.detail[NumTob].PutValue('GEX_VALDEFAUT','-');
        end
      else if P.Controls[j] is THEdit then
        Tob_Champs.detail[NumTob].PutValue('GEX_VALDEFAUT',THEdit(P.Controls[j]).Text)
      else if P.Controls[j] is THValComboBox then
        Tob_Champs.detail[NumTob].PutValue('GEX_VALDEFAUT',THValComboBox(P.Controls[j]).value);
      end;
    end;
end;

end.
