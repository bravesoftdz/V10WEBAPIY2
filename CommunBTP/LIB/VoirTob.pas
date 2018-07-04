unit VoirTob;

interface

uses  Classes, ComCtrls, StdCtrls, SysUtils, HCtrls,
			variants, 
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      FE_Main,
{$ENDIF}
      AGLInit, UTOF, Utob, HPanel, Vierge;

Type

     TOF_VOIRTOB = Class (TOF)
     GS : THGrid;
     TV : TTreeView;
     IM : TCheckBox;
     PGS : THPanel;
   private
     CLengths : array [0..3] of integer;
     LesChamps : string;
  	 WithChamps,WithChampsSup,WithTypes : boolean;
     procedure TVChange(Sender: TObject; Node: TTreeNode);
     procedure IMClick(Sender: TObject);
     procedure PGSResize(Sender: TObject);
		 function DecodeType (Donnee : variant) : string;
   public
     procedure OnLoad ; override ;
     procedure Onargument ( Arguments : string) ; override ;
   END ;

procedure GCVoirTob ( ToTob : Tob ; Titre : string = '' ; lesChamps : string = ''; Champs: boolean = true;ChampsSup : boolean = true;AvecType : boolean=false);

implementation

procedure GCVoirTob ( ToTob : Tob ; Titre : string = '' ; lesChamps : string = ''; Champs: boolean = true;ChampsSup : boolean = true;AvecType : boolean=false);
var TOBTitre : TOB;
begin
TheTob := ToTob;
//if (Titre <> '') or (LesChamps <> '') then
  begin
  TOBTitre := TOB.create ('LE TITRE',nil,-1);
  TOBTitre.addChampsupValeur ('TITRE',Titre,false);
  TOBTitre.AddChampSupValeur ('LESCHAMPS',LesChamps);
  TOBTitre.AddChampSupValeur ('WITHCHAMPS',Champs);
  TOBTitre.AddChampSupValeur ('WITHCHAMPSSUP',ChampsSup);
  TOBTitre.AddChampSupValeur ('WITHTYPE',AvecType);
  TheTob.data :=TobTitre ;
  end;
AGLLanceFiche('GC','GCVOIRTOB','','','');
if titre <> '' then TOBTitre.free;
end ;

Procedure TOF_VOIRTOB.Onload ;
var
    ind1 : integer;
    st1 : string;

    procedure AjouteChamp(TobATraiter : TOB);
    var  TT1 : TOB;
         ind1 : integer;
    begin
(*    if not TobATraiter.FieldExists('NOMTAB') then
        TobATraiter.AddChampSup('NOMTAB', False);
    if TobATraiter.NomTable <> '' then
        TobATraiter.PutValue('NOMTAB', TobATraiter.NomTable)
        else
        TobATraiter.PutValue('NOMTAB', 'Indéfini');
*)
    for ind1 := 0 to TobATraiter.Detail.Count -1 do
        begin
        TT1 := TobATraiter.Detail[ind1];
        AjouteChamp(TT1);
        end;
    end;

begin
inherited ;
    GS := THGrid(GetControl('GS'));
    if WithTypes then GS.ColCount := 4
    						 else GS.colCount := 3;
    CLengths[0] := Trunc((GS.ColWidths[0] / GS.ClientWidth) * 100);
    CLengths[1] := Trunc((GS.ColWidths[1] / GS.ClientWidth) * 100);
    CLengths[2] := Trunc((GS.ColWidths[2] / GS.ClientWidth) * 100);
    CLengths[3] := Trunc((GS.ColWidths[3] / GS.ClientWidth) * 100);

    GS.ColTypes[2] := 'B';
    GS.ColFormats[2] := IntToStr(Ord(csCheckBox));
    GS.ColAligns[2] := taCenter;

    TV := TTreeView(GetControl('TV'));
    TV.OnChange := TVChange;
    IM := TCheckBox(GetControl('IM'));
    IM.OnClick := IMClick;
    PGS := THPanel(GetControl('PGS'));
    PGS.OnResize := PGSResize;
    AjouteChamp(LaTob);
(*
    for ind1 := 0 to LaTob.MaxNiveau do
        st1 := st1 + 'NOMTAB;';
*)
    LaTob.PutTreeView(TV, nil, st1);
UpdateCaption (ecran);
// LaTob.PutGridDetail(THGRID(GetControl('GS')),True,True,'',True);
end;

procedure TOF_VOIRTOB.TVChange(Sender: TObject; Node: TTreeNode);
var Tobassoc : TOB;
    ind1, ind2 : integer;
    NomChamp : string;
begin
  TobAssoc := TOB(Node.Data);
  for ind1 := 1 to GS.RowCount do GS.Rows[ind1].Clear;
  ind1 := 1;
  ind2 := 1;
  if WithChamps then
  begin
    NomChamp := TobAssoc.GetNomChamp(ind2);
    while NomChamp <> '' do
    begin
      if (Pos(NomChamp,lesChamps) > 0) or (lesChamps = '') then
      begin
        if ind1 > 2 then GS.RowCount := ind1;
        if (TobAssoc.IsFieldModified(NomChamp) and IM.Checked) or (not IM.Checked) then
        begin
          GS.Cells[0, ind1] := NomChamp;
          if (VarIsNull (TobAssoc.GetValeur(ind2))) or (VarAsType(TobAssoc.GetValeur(ind2), varString) = #0) then
          	GS.cells[1,ind1] := '(NULL)'
          else
          	GS.Cells[1, ind1] := VarAsType(TobAssoc.GetValeur(ind2), varString);
          if TobAssoc.IsFieldModified(NomChamp) then GS.Cells[2, ind1] := 'X'
                                                else GS.Cells[2, ind1] := '-';
          if WithTypes then
          begin
          	GS.cells[3,Ind1] := DecodeType (TOBAssoc.GetValeur(Ind2));
          end;
          Inc(ind1);
        end;
      end;
      Inc(ind2);
      NomChamp := TobAssoc.GetNomChamp(ind2);
    end;
  end;
	if WithChampsSup then
  begin
    ind2 := 1000;
    NomChamp := TobAssoc.GetNomChamp(ind2);
    while NomChamp <> '' do
    begin
      if (Pos(NomChamp,lesChamps) > 0) or (lesChamps = '') then
      begin
        GS.RowCount := ind1 + 1;
        if (TobAssoc.IsFieldModified(NomChamp) and IM.Checked) or (not IM.Checked) then
        begin
          GS.Cells[0, ind1] := NomChamp;
          if (VarIsNull (TobAssoc.GetValeur(ind2))) or (VarAsType(TobAssoc.GetValeur(ind2), varString) = #0) then
          	GS.cells[1,ind1] := '(NULL)'
          else
          	GS.Cells[1, ind1] := VarAsType(TobAssoc.GetValeur(ind2), varString);
          if TobAssoc.IsFieldModified(NomChamp) then GS.Cells[2, ind1] := 'X'
                                                else GS.Cells[2, ind1] := '-';
          if WithTypes then
          begin
          	GS.cells[3,Ind1] := DecodeType (TOBAssoc.GetValeur(Ind2));
          end;
          Inc(ind1);
        end;
      end;
      Inc(ind2);
      NomChamp := TobAssoc.GetNomChamp(ind2);
    end;
  end;
end;

function TOF_VOIRTOB.DecodeType (Donnee : variant) : string;
var TheType : integer;
begin
	TheType :=  VarType (Donnee);
  result := 'NON DEFINI';
  if (TheType = VarInteger) or (TheType = VarSmallInt) then Result := 'INTEGER'
  else if (TheType = VarDouble) or (TheTYpe = VarCurrency) then result := 'DOUBLE'
  else if (TheTYpe = VarBoolean) then result := 'BOOLEAN'
  else if (TheTYpe = VarString) then result := 'STRING'
  else if (TheType = VarEmpty) then result := 'NULL';

end;

procedure TOF_VOIRTOB.IMClick(Sender: TObject);
begin
TVChange(Sender, TV.Selected);
end;

procedure TOF_VOIRTOB.PGSResize(Sender: TObject);
begin
GS.ColWidths[0] := Trunc((GS.ClientWidth / 100) * CLengths[0]);
GS.ColWidths[1] := Trunc((GS.ClientWidth / 100) * CLengths[1]);
GS.ColWidths[2] := Trunc((GS.ClientWidth / 100) * CLengths[2]);
end;


procedure TOF_VOIRTOB.Onargument(Arguments: string);
begin
  inherited;
  if laTob.data <> nil then
  begin
  	if TOB(laTob.data).getValue('TITRE') <> '' then ecran.caption := TOB(laTob.data).getValue('TITRE');
		if TOB(laTob.data).getValue('LESCHAMPS') <> '' then LesChamps := TOB(laTob.data).getValue('LESCHAMPS');
  	WithChamps :=  TOB(laTob.data).getValue('WITHCHAMPS');
    WithChampsSup := TOB(laTob.data).getValue ('WITHCHAMPSSUP');
  	WithTypes := TOB(laTob.data).getValue ('WITHTYPE');
  end else
  begin
  	lesChamps := '';
  	WithChamps := true;
    WithChampsSup := True;
  	WithTypes :=false;
  end;
end;

Initialization
registerclasses([TOF_VOIRTOB]);
end.
