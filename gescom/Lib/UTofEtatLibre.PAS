{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 23/03/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE :
Mots clefs ... : TOF;ETATLIBRE
*****************************************************************}
Unit UTofEtatLibre ;

Interface

Uses StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,
{$ELSE}
     db,dbTables,QRS1,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,UTOB,M3FP,Filtre;

Type
  TOF_ETATLIBRE = Class (TOF)
    public
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;

    procedure TableEnter(Sender: TObject);
    procedure TableClick(Sender: TObject);
    procedure ChampClick(Sender: TObject);
    procedure ChampEnter(Sender: TObject);
    procedure BoxClick(Sender: TObject);
    procedure FiltreChange(Sender: TObject);
    procedure TriTableClick(Sender: TObject);

    private
    Champ_Text : string;
    OTable_Text,OTable_Value : string;
    TobChamp : TOB;
    //initialisation
    procedure InitChamp(stNumTable : string);
    procedure TriTable(NomChamp : string);
    //Maj des données des champs combos liées à une table
    procedure MAJChamp(stNomChamp : string );
    procedure MAJChampTri(NewTable_Value,NewTable_Text : string);
    //Evénments sur TableClick
    procedure ModifValCombo(stTable,stNumC : string);
    procedure ChangePropriete(stNomChamp : string);
    //Gestion de TobChamp
    procedure TobAjout(stTable : string);
    //Sur FiltreChange
    procedure ChargeValeur(stNomChamp,stVal : string);

    function  TraduitOperateur ( stOpe, stVal : String ) : String ;
    function  ExtraitNumChamp(stNomChamp : string) : string;

  end ;


Implementation

procedure TOF_ETATLIBRE.OnArgument (StArgument : String) ;
var i_ind : integer;
    CChamp, CFiltre : THValComboBox;
begin
  Inherited ;
Champ_Text := '';
OTable_Text := '';
OTable_Value:= '';
CChamp := nil;
TobChamp := TOB.Create('',nil,-1);
/////Add des items des Combos liés aux tables
for i_ind := 0 to Ecran.ComponentCount -1 do
    begin
    if Ecran.Components[i_ind] is THValComboBox then
     CChamp := THValComboBox(Ecran.Components[i_ind]) else continue;
    if (pos('CHAMP',CChamp.Name) > 0) or (pos('FEtat',CChamp.Name) > 0) then
       begin
       CChamp.OnClick := ChampClick;
       CChamp.OnEnter := ChampEnter;
       end;
    if pos('_TABLE',CChamp.Name) > 0 then
       begin
       CChamp.OnEnter := TableEnter;
       CChamp.OnClick := TableClick;
       TriTable(CChamp.Name);
       end;
    if pos('_TRITABLE',CChamp.Name) > 0 then CChamp.OnClick := TriTableClick;
    end;
CFiltre := THValComboBox(GetControl('FFILTRES'));
CFiltre.OnChange := FiltreChange;
if CFiltre.Text <> '' then FiltreChange(CFiltre);
end ;

procedure TOF_ETATLIBRE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_ETATLIBRE.OnLoad ;
var i_ind : integer;
    CChamp : THValComboBox;
    stWhere,stOrder,stComp,stVal,stCoord : string;
    stNumC : string;
begin
  Inherited ;
//if GetControlText('_TABLE') = '' then exit;
stWhere := ''; stOrder := ''; stCoord := '';   CChamp := nil;
SetControlText('XX_WHERE','');
SetControlText('XX_ORDERBY','');
/////Construction du Where
for i_ind := 0 to  Ecran.ComponentCount -1 do
    begin
    if Ecran.Components[i_ind] is THValComboBox then
     CChamp := THValComboBox(Ecran.Components[i_ind]) else continue;
    if (pos('_CHAMP',CChamp.Name) > 0) and (CChamp.Value <> '') then
       begin
       stNumC := ExtraitNumChamp(CChamp.Name);
       stComp := GetControlText('_COMP'+stNumC);
       if GetControl('_CB'+stNumC)<> nil then
            stVal := GetControlText('_CB'+stNumC)
       else begin
            stVal := GetControlText('_VAL'+stNumC);
            if THEdit(GetControl('_VAL'+stNumC)).OpeType=otDate then
               stVal := USDATETIME(strToDate(stVal));
            end;
       if (stComp <> '') and (stVal <> '') then
          begin
          stWhere := stWhere + stCoord + CChamp.Value + TraduitOperateur(stComp,stVal);
          stCoord := ' ' + GetControlText('_COORD'+stNumC) + ' ';
          if Trim(stCoord) = '' then stCoord := ' AND '; //par défaut
          end;
       end;
    if (pos('_TRICHAMP',CChamp.Name) > 0) and (CChamp.Value <> '') then
       begin
       if stOrder = '' then stOrder := CChamp.Value
       else stOrder := stOrder + ',' + CChamp.Value;
       end;
    end;
SetControlText('XX_WHERE', stWhere);

/////Order
if stOrder <> '' then
   begin
   if GetControlText('_CTRIDESC')= 'X' then stOrder := stOrder + ' DESC'
   else stOrder := stOrder + ' ASC';
   end;
SetControlText('XX_ORDERBY', stOrder);
end ;

procedure TOF_ETATLIBRE.OnClose ;
begin
  Inherited ;
TobChamp.Free;
end ;

////////////////////////////////////////////////////////////////////////////////
/////Gestion des champs
procedure TOF_ETATLIBRE.TableEnter(Sender: TObject);
begin
OTable_Text := THValComboBox(Sender).Text;
OTable_Value := THValComboBox(Sender).Value;
end;

procedure TOF_ETATLIBRE.TableClick(Sender: TObject);
var CListeTable : THValComboBox;
TobC : TOB;
    i_index : integer;
    stNumT : string;
begin
if THValComboBox(Sender).Value = OTable_Value then exit;
CListeTable := THValComboBox(GetControl('_LISTETABLE'));
i_index := CListeTable.Items.IndexOf(OTable_Value);
if i_index > -1 then CListeTable.items.Delete(i_index);
CListeTable.Items.Add(THValComboBox(Sender).Value);

stNumT := ExtraitNumChamp(THValComboBox(Sender).Name);
InitChamp(stNumT); //Réinitialisation des champs
if CListeTable.Items.IndexOf(OTable_Value) = -1 then //si table pa déjà existante
   begin
   TobC := TobChamp.FindFirst(['TABLE'],[OTable_Value],false);
   if TobC <> nil then TobC.Free;
   end;
TobAjout(THValComboBox(Sender).Value);
MAJChamp(THValComboBox(Sender).Name);
MAJChampTri(THValComboBox(Sender).Value,THValComboBox(Sender).Text);
OTable_Text := THValComboBox(Sender).Text;
OTable_Value := THValComboBox(Sender).Value;
end;

procedure TOF_ETATLIBRE.ChampClick(Sender: TObject);
var CChamp : THValComboBox;
    stNumC : string;
begin
CChamp := THValComboBox(Sender);
if pos('FEtat',CChamp.Name) >0 then
   begin
   SetControlText('_CFETAT',CChamp.value);
   TFQRS1(Ecran).CodeEtat := GetControlText('_CFETAT');
   exit;
   end;
stNumC := ExtraitNumChamp(CChamp.Name);
if pos('TRICHAMP',CChamp.Name) >0 then
   begin
   SetControlText('_CTRICHAMP'+stNumC,CChamp.value);
   exit;
   end;
SetControlText('_C'+stNumC,CChamp.value);
if CChamp.Value <> Champ_Text then ChangePropriete(CChamp.Name);
Champ_Text := CChamp.Value;
end;

procedure TOF_ETATLIBRE.ChampEnter(Sender: TObject);
var CChamp : THValComboBox;
begin
CChamp := THValComboBox(Sender);
Champ_Text := CChamp.Value;
end;

procedure TOF_ETATLIBRE.BoxClick(Sender: TObject);
var CVal : THEdit;
begin
CVal := THEdit(GetControl('_VAL'+ExtraitNumChamp(TCheckBox(Sender).Name)));
if CVal = nil then exit;
if TCheckBox(Sender).Checked then CVal.Text := 'X' else CVal.Text := '-';
end;

procedure TOF_ETATLIBRE.FiltreChange(Sender: TObject);
var F : TForm;
    i_ind : integer;
    CChamp,CListeTable : THValComboBox;
    stChamp,stNum,stVal : string;
begin
F:=TForm(Ecran);
CListeTable := THValComboBox(GetControl('_LISTETABLE'));
if TFQRS1(F).FNomFiltre = '' then exit;
OTable_Value := '' ; OTable_Text := '';
TobChamp.ClearDetail;
CListeTable.Items.Clear;
for i_ind := Ecran.ComponentCount -1 downto 0 do
   if Ecran.Components[i_ind] is THValComboBox then
      begin
      if pos('_TABLE',THValComboBox(Ecran.Components[i_ind]).Name)>0 then
         InitChamp(ExtraitNumChamp(THValComboBox(Ecran.Components[i_ind]).Name));
      if pos('_TRI',THValComboBox(Ecran.Components[i_ind]).Name)>0 then
         begin
         CChamp := THValComboBox(GetControl(Ecran.Components[i_ind].Name));
         CChamp.Items.Clear; CChamp.Values.Clear;
         end;
      end;

with TFQRS1(F) do FFiltresChange(nil);
CChamp := THValComboBox(GetControl('FEtat'));
CChamp.ItemIndex := CChamp.Values.IndexOf(GetControlText('_CFETAT'));
TFQRS1(F).CodeEtat := GetControlText('_CFETAT');
for i_ind := 0 to Ecran.ComponentCount -1 do
    begin
    if Ecran.Components[i_ind] is THValComboBox then
    CChamp := THValComboBox(Ecran.Components[i_ind]) else continue;
    if pos('_TABLE',CChamp.Name) > 0 then
       begin
       CListeTable.Items.Add(CChamp.Value);
       TobAjout(CChamp.Value);
       MAJChamp(CChamp.Name);
       MAJChampTri(CChamp.Value,CChamp.Text);
       end;
    end;
for i_ind := Ecran.ComponentCount -1 downto 0 do 
    begin
    if Ecran.Components[i_ind] is THValComboBox then
    CChamp := THValComboBox(Ecran.Components[i_ind]) else continue;
    stChamp := '';
    if (pos('_CHAMP',CChamp.Name) <= 0) and (pos('_TRI',CChamp.Name) <= 0) then continue;
    stNum := ExtraitNumChamp(CChamp.Name);
    if pos('_CHAMP',CChamp.Name) > 0 then
       begin
       stChamp := GetControlText('_C'+stNum);
       if stChamp <> '' then
          begin
          CChamp.ItemIndex := CChamp.Values.IndexOf(stChamp);
          stVal := GetControlText('_VAL'+stNum);
          ChangePropriete(CChamp.Name);
          ChargeValeur(CChamp.Name,stVal);
          end;
       end else
       begin
       if pos('_TRITABLE',CChamp.Name) > 0 then stChamp := GetControlText('_CTRITABLE'+stNum)
       else if pos('_TRICHAMP',CChamp.Name) > 0 then
            begin
            stChamp := GetControlText('_CTRICHAMP'+stNum);
            ModifValCombo(GetControlText('_CTRITABLE'+stNum), stNum);
            end;
       if stChamp <> '' then CChamp.ItemIndex := CChamp.Values.IndexOf(stChamp);
       end;
    end;
end;

procedure TOF_ETATLIBRE.TriTableClick(Sender: TObject);
var stNumC : string;
begin
stNumC := ExtraitNumChamp(THValComboBox(Sender).Name);
SetControlText('_CTRITABLE'+stNumC,THValComboBox(Sender).value);
ModifValCombo(THValComboBox(Sender).Value,stNumC);
end;

///////////////////////////////////////////////////////////////////////////////
procedure TOF_ETATLIBRE.InitChamp(stNumTable : string);
var i_ind,i_ind2 : integer;
    EVal : THEdit;
    CB : TCheckBox;
    COpe : THValComboBox;
    GB : TGroupBox;
begin
GB := nil;
for i_ind := Ecran.ComponentCount -1 downto 0 do
    begin
    if Ecran.Components[i_ind] is TGroupBox then
       begin
       GB := TGroupBox(Ecran.Components[i_ind]);
       if ExtraitNumChamp(GB.Name) <> stNumTable then continue;
       end else continue;
    for i_ind2 := 0 to GB.ControlCount -1 do
        begin
        if GB.Controls[i_ind2] is THEdit then
           begin
           EVal := THEdit(GB.Controls[i_ind2]);
           if (pos('_VAL',EVal.Name) > 0) then
              begin
              EVal.Visible := True;
              EVal.Text := '';
              EVal.DataType := '';
              EVal.ElipsisButton := False;
              EVal.OpeType := otString;
              EVal.EditMask := '';
              end;
           end else
           if GB.Controls[i_ind2] is TCheckBox then
              begin
              CB := TCheckBox(GB.Controls[i_ind2]);
              if (pos('_CB',CB.Name) > 0) then begin CB.Free ; continue ; end;
              end;
           if GB.Controls[i_ind2] is THValComboBox then
              begin
              COpe := THValComboBox(GB.Controls[i_ind2]);
              if (pos('_COMP',COpe.Name) > 0)  then COpe.ItemIndex := -1;
              if (pos('_COORD',COpe.Name) > 0) then COpe.ItemIndex := -1;
              end;
        end;
    end;
end;

procedure TOF_ETATLIBRE.TriTable(NomChamp : string);
begin
if ctxGescom in V_PGI.PGIContexte then
   SetControlProperty(NomChamp,'PLUS','DT_DOMAINE="G" OR DT_DOMAINE="Y" OR DT_DOMAINE="A"');
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ETATLIBRE.MAJChamp(stNomChamp : string);
var CChamp : THValComboBox ;
    GB : TGroupBox;
    TobC : TOB;
    i_ind1,i_ind2,i_ind3 : integer;
    stNumT, Table : string;
begin
Table  := GetControlText(stNomChamp);
stNumT := ExtraitNumChamp(stNomChamp);
TobC := TobChamp.FindFirst(['TABLE'],[Table],false);
if TobC = nil then exit;
CChamp := nil ; GB := nil;
for i_ind1 := 0 to  Ecran.ComponentCount -1 do
  begin
  if Ecran.Components[i_ind1] is TGroupBox then
     begin
     GB := TGroupBox(Ecran.Components[i_ind1]);
     if ExtraitNumChamp(GB.Name) <> stNumT then continue;
     end else continue;
  for i_ind2 := 0 to GB.ControlCount -1 do
      begin
      if GB.Controls[i_ind2] is THValComboBox then
      CChamp := THValComboBox(GB.Controls[i_ind2]) else continue;
      if pos('_CHAMP',CChamp.Name) > 0 then
         begin
         CChamp.Items.Clear; CChamp.Values.Clear;
         CChamp.Items.Add('<<Aucun>>');
         CChamp.Values.Add('');
         for i_ind3 := 0 to TobC.Detail.Count -1 do
             begin
             CChamp.Values.Add(TobC.Detail[i_ind3].GetValue('DH_NOMCHAMP'));
             CChamp.Items.Add(TobC.Detail[i_ind3].GetValue('DH_LIBELLE'));
             end;
         end;
      end;
  end;
end;

procedure TOF_ETATLIBRE.MAJChampTri(NewTable_Value,NewTable_Text : string);
var i_ind, i_index : integer;
    stNumC : string;
    CCTable, CChamp, CListeTable : THValComboBox;
begin
CCTable := nil ;
CListeTable := THValComboBox(GetControl('_LISTETABLE'));
for i_ind := 0 to Ecran.ComponentCount -1 do
    begin
    if Ecran.Components[i_ind] is THValComboBox then
       CCTable := THValComboBox(Ecran.Components[i_ind]) else continue;
    if pos('_TRITABLE',CCTable.Name) > 0 then
       begin
       if (NewTable_Value <>'') and (CCTable.Values.IndexOf(NewTable_Value) = -1) then
           begin
           CCTable.Values.Add(NewTable_Value);
           CCTable.Items.Add(NewTable_Text);
           end;
       i_index := CCTable.Values.IndexOf(OTable_Value);
       if (OTable_Value <> '') and (i_index > -1) and (CListeTable.Values.IndexOf(OTable_Value) = -1) then
           begin
           stNumC := ExtraitNumChamp(CCTable.Name);
           if CCTable.Value = OTable_Value then
              begin
              CChamp := THValComboBox(GetControl('_TRICHAMP'+stNumC));
              CChamp.Items.Clear; CChamp.Values.Clear;
              CChamp.Items.Add('<<Aucun>>');
              CChamp.Values.Add('');
              THEdit(GetControl('_CTRICHAMP'+stNumC)).Text:='';
              end;
           //if CCTable.Value = OTable_Value then
           CCTable.items.Delete(i_index);
           CCTable.Values.Delete(i_index);
           THEdit(GetControl('_CTRITABLE'+stNumC)).Text:='';
           end;
       end;
   end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ETATLIBRE.ModifValCombo(stTable, stNumC : string);
var i_ind : integer;
    CChamp : THValComboBox ;
    TobC : TOB;
begin
if stTable = '' then exit;
TobC := TobChamp.FindFirst(['TABLE'],[stTable],false);
if TobC = nil then exit;
CChamp := THValComboBox(GetControl('_TRICHAMP'+stNumC));
CChamp.Items.Clear; CChamp.Values.Clear;
CChamp.Items.Add('<<Aucun>>');
CChamp.Values.Add('');
for i_ind := 0 to TobC.Detail.Count -1 do
    begin
    CChamp.Values.Add(TobC.Detail[i_ind].GetValue('DH_NOMCHAMP'));
    CChamp.Items.Add(TobC.Detail[i_ind].GetValue('DH_LIBELLE'));
    end;
end;

procedure TOF_ETATLIBRE.ChangePropriete(stNomChamp : string);
var TobC, TobT : TOB;
    stType,stNumC, Table : string;
    EVal : THEdit;
    CB : TCheckBox;
    CTable : THValComboBox;
    GB : TGroupBox;
begin
stType := '';
stNumC := ExtraitNumChamp(stNomChamp);

EVal := THEdit(GetControl('_VAL'+stNumC));
EVal.ElipsisButton := False;
EVal.EditMask := '';
EVal.Text := '';
TobC := nil ;
GB := TGroupBox(GetControl(stNomChamp).GetParentComponent);
CTable := THValComboBox(GetControl('_TABLE'+ExtraitNumChamp(GB.Name)));
Table := CTable.Value;
TobT := TobChamp.FindFirst(['TABLE'],[Table],false);
if TobT = nil then exit;
if TobT.Detail.Count > 0 then
   TobC := TobT.FindFirst(['DH_NOMCHAMP'],[GetControlText(stNomChamp)],false);
if TobC <> nil then stType := TobC.GetValue('DH_TYPECHAMP') else exit;
CB := TCheckBox(GetControl('_CB'+stNumC));
if pos('BOOLEAN',stType)>0 then
    begin
    if CB = nil then
       begin
       EVal.Visible := false;
       CB:=TCheckBox.Create(Ecran); CB.Parent:=EVal.Parent; CB.Name:='_CB'+stNumC;
       CB.Left:=EVal.Left ; CB.Width:=30  ; CB.Top := EVal.Top ;
       CB.Caption:=''; CB.TabStop := true ; CB.TabOrder := EVal.TabOrder;CB.OnClick := BoxClick;
       end else CB.Checked:=false;
    exit;
    end else
    begin
    if CB <> nil then
       begin
       CB.Free;
       EVal.Visible := true;
       end;
    end;
if (pos('DOUBLE',stType)>0) or (pos('INTEGER',stType)>0) then EVal.OpeType :=otReel
else if pos('DATE',stType)>0 then
        begin
        EVal.OpeType := otDate;
        EVal.EditMask:='!99/99/0000;1;_';
        EVal.ControlerDate := true;
        end else EVal.OpeType := otString;
EVal.DataType := Get_Join(GetControlText(stNomChamp));
if EVal.DataType <> '' then EVal.ElipsisButton := True;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ETATLIBRE.TobAjout(stTable : string);
var QChamp : TQuery;
    TobC : TOB;
begin
TobC := TobChamp.FindFirst(['TABLE'],[stTable],false);
if TobC <> nil then exit;
TobC := TOB.Create('',TobChamp,-1);
TobC.AddChampSup('TABLE',false);
TobC.PutValue('TABLE',stTable);
//Recup des champs de la table
QChamp := OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE,DH_TYPECHAMP FROM DECHAMPS WHERE DH_PREFIXE="'+
            TableToPrefixe(stTable)+'" AND DH_CONTROLE LIKE "%D%"',True) ;

if not QChamp.Eof then TobC.LoadDetailDb('DECHAMPS','','',QChamp,false);
Ferme(QChamp);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ETATLIBRE.ChargeValeur(stNomChamp, stVal : string);
var CVal : THEdit;
    CTable : THValComboBox;
    GB : TGroupBox;
    TypeC, stNumC, Table : string;
    TobT, TobC : TOB;
begin
stNumC := ExtraitNumChamp(stNomChamp);
CVal := THEdit(GetControl('_VAL'+stNumC));
if CVal = nil then exit;
TobC := nil ;
GB := TGroupBox(GetControl(stNomChamp).GetParentComponent);
CTable := THValComboBox(GetControl('_TABLE'+ExtraitNumChamp(GB.Name)));
Table := CTable.Value;
TobT := TobChamp.FindFirst(['TABLE'],[Table],false);
if TobT = nil then exit;
if TobT.Detail.Count > 0 then
   TobC := TobT.FindFirst(['DH_NOMCHAMP'],[GetControlText('_CHAMP'+stNumC)],false);
if TobC <> nil then TypeC := TobC.GetValue('DH_TYPECHAMP') else exit;
if pos('BOOLEAN',TypeC) > 0 then
   begin
   TCheckBox(GetControl('_CB'+stNumC)).Checked := (stVal = 'X');
   end else CVal.Text := stVal;
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_ETATLIBRE.ExtraitNumChamp(stNomChamp : string) : string;
var ind : integer;
begin
for ind := length(stNomChamp) downto 0 do
    if not isNumeric(stNomChamp[ind]) then break;
result := copy(stNomChamp,ind+1,length(stNomChamp));
end;

/////Traduit la tablette TTCompare
function TOF_ETATLIBRE.TraduitOperateur ( stOpe, stVal : String ) : String ;
begin
if stOpe = 'C' then    // commence par
   begin
   stOpe := ' LIKE ' ;
   stVal := stVal + '%' ;
   end else
if stOpe = 'D' then    // ne commence pas par
   begin
   stOpe := ' NOT LIKE ' ;
   stVal :=  stVal + '%';
   end else
if stOpe = 'L' then    // contient
   begin
   stOpe := ' LIKE ' ;
   stVal := '%' + stVal + '%' ;
   end else
if stOpe = 'M' then    // ne contient pas
   begin
   stOpe := ' NOT LIKE ' ;
   stVal := '%' + stVal + '%' ;
   end ;
Result := stOpe + '"' + stVal + '"' ;
end ;


Initialization
  registerclasses ( [ TOF_ETATLIBRE ] ) ;
  
end.


