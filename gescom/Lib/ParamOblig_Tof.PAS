{***********UNITE*************************************************
Auteur  ...... : BBY
Créé le ...... : 04/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARAMOBLIG ()
Mots clefs ... : TOF;PARAMOBLIG
*****************************************************************}
Unit PARAMOBLIG_TOF ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTb97,
     UTOB,UTOF ;

type
  TListeChps = Class
   Champ     : String ;
   Intitule  : String ;
   END ;

Type
  TOF_PARAMOBLIG = Class (TOF)
    LB1Libelle: TListBox;
    LB2Libelle: TListBox;
    bSel: TToolbarButton97;
    bSelAll: TToolbarButton97;
    bDeSel: TToolbarButton97;
    bDeSelAll: TToolbarButton97;
    bAnnule: TToolbarButton97;
    bValide: TToolbarButton97;
    LB1Nom: TListBox;
    LB2Nom: TListBox;
    RG1: THRadioGroup;
    RG2: THRadioGroup;
    HValCB1: THValComboBox;
    TableActive: THEdit;
    CodeActif: THEdit;
    CodePrec: THEdit;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    bModif : boolean;
    bModifFin : boolean;
    TobFic : TOB;
    procedure HValCB1Click(Sender: TObject);
    procedure LB1LibelleDblClick(Sender: TObject);
    procedure LB2LibelleClick(Sender: TObject);
    procedure LB2NomClick(Sender: TObject);
    procedure LB2LibelleDblClick(Sender: TObject);
    procedure LB2NomDblClick(Sender: TObject);
    procedure bSelClick(Sender: TObject);
    procedure bDeSelClick(Sender: TObject);
    procedure bDeSelAllClick(Sender: TObject);
    procedure bAnnuleClick(Sender: TObject);
    procedure bValideClick(Sender: TObject);
    procedure RG1Click(Sender: TObject);
    procedure RG2Click(Sender: TObject);
    procedure LitDonnees;
    procedure EnableButtons;
  end ;

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Voulez-vous enregistrer les modifications ?',
          {2}  'Vos modifications seront actives à la prochaine connexion.'
              );

Implementation
uses 
   CbpMCD
   ,CbpEnumerator
;

procedure TOF_PARAMOBLIG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARAMOBLIG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARAMOBLIG.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PARAMOBLIG.OnLoad ;
begin
    Inherited ;
    LB1Libelle := TListBox(GetControl('LB1LIBELLE'));
    LB1Libelle.OnDblClick := LB1LibelleDblClick;
    LB2Libelle := TListBox(GetControl('LB2LIBELLE'));
    LB2Libelle.OnClick := LB2LibelleClick;
    LB2Libelle.OnDblClick := LB2LibelleDblClick;
    bSel := TToolbarButton97(GetControl('BSEL'));
    bSel.OnClick := bSelClick;
    bSelAll := TToolbarButton97(GetControl('BSELALL'));
    bSelALL.Visible := False;
    bDeSel := TToolbarButton97(GetControl('BDESEL'));
    bDeSel.OnClick := bDeSelClick;
    bDeSelAll := TToolbarButton97(GetControl('BDESELALL'));
    bDeSelAll.OnClick := bDeSelAllClick;
    bAnnule := TToolbarButton97(GetControl('BFERME'));
    bAnnule.OnClick := bAnnuleClick;
    bAnnule.Visible := True;
    bValide := TToolbarButton97(GetControl('BVALIDER'));
    bValide.OnClick := bValideClick;
    LB1Nom := TListBox(GetControl('LB1NOM'));
    LB2Nom := TListBox(GetControl('LB2NOM'));
    LB2Nom.OnClick := LB2NomClick;
    LB2Nom.OnDblClick := LB2NomDblClick;
    RG1 := THRadioGroup(GetControl('RG1'));
    RG1.OnClick := RG1Click;
    RG2 := THRadioGroup(GetControl('RG2'));
    RG2.OnClick := RG2Click;
    HValCB1 := THValComboBox(GetControl('HVALCB1'));
    HValCB1.OnClick := HValCB1Click;
    TableActive := THEdit(GetControl('TABLEACTIVE'));
    CodeActif := THEdit(GetControl('CODEACTIF'));
    CodePrec := THEdit(GetControl('CODEPREC'));
    TableActive.Text := '';
    CodeActif.Text := '';
    CodePrec.Text := '';
    bModif := False;
    bModifFin := False;
    TobFic := TOB.Create('', nil, -1);
    LB1Nom.Top := LB1Libelle.Top;
    LB1Nom.Left := LB1Libelle.Left;
    LB1Nom.Width := LB1Libelle.Width;
    LB1Nom.Height := LB1Libelle.Height;
    LB2Nom.Top := LB2Libelle.Top;
    LB2Nom.Left := LB2Libelle.Left;
    LB2Nom.Width := LB2Libelle.Width;
    LB2Nom.Height := LB2Libelle.Height;
    LB1Libelle.Sorted:=False;
    LB1Nom.Sorted:=False;
    LB2Libelle.Sorted:=False;
    LB2Nom.Sorted:=False;
    RG1Click(Self);
    RG2Click(Self);
    EnableButtons;
end ;

procedure TOF_PARAMOBLIG.OnArgument (S : String ) ;
var Combo : THValComboBox;
    streq : String;
begin
  Inherited ;
            // mcd 22/05/02 pour prendre en compte les domaines
  Combo:=THValComboBox(GetControl('HVALCB1'));

  // gm le 06/10/03
  if (ctxaffaire in V_PGI.PGIContexte) then
    StReq :='AND (CO_CODE like "A%" OR CO_CODE like "B%" '
  else
    if (ctxgcaff in V_PGI.PGIContexte) then
      StReq:='AND (CO_CODE like "G%" OR CO_CODE like "A%" '
    else
      if (ctxChr in V_PGI.PGIContexte) then
        StReq:='AND (CO_CODE like "G%" OR CO_CODE like "H%" '
      else
        StReq:='AND (CO_CODE like "G%"  ';

  {$IFDEF GRC} { on ne prend pas les infs complémentaires qui sont gérés à part pour les champs obligatoires }
  if (ctxGRC in V_PGI.PGIContexte) then
    StReq := StReq + ' OR ( (CO_CODE like "R%") AND (CO_CODE <> "R02") AND (CO_CODE <> "R03")) )'
  else
    StReq :=  StReq + ')';
  {$ELSE}
    StReq :=  StReq + ')';
  {$ENDIF}
  Combo.Plus := StReq;
  Combo.ExHaustif:=Explus;
end ;

procedure TOF_PARAMOBLIG.OnClose ;
Var i_ind : integer ;
begin
  Inherited ;
if bModifFin then
  PGIBox(TexteMessage[2], Ecran.Caption);
TOBFic.free;
for i_ind:=0 to LB1Libelle.Items.Count-1 do
   TListeChps(LB1Libelle.Items.Objects[i_ind]).Free ;
for i_ind:=0 to LB2Libelle.Items.Count-1 do
   TListeChps(LB2Libelle.Items.Objects[i_ind]).Free ;
end ;

procedure TOF_PARAMOBLIG.HValCB1Click(Sender: TObject);
begin
CodeActif.Text := HValCB1.Value;
if CodePrec.Text = '' then CodePrec.Text := CodeActif.Text;
LitDonnees;
end;

procedure TOF_PARAMOBLIG.LB1LibelleDblClick(Sender: TObject);
begin
bSelClick(Nil);
end;

procedure TOF_PARAMOBLIG.LB2LibelleClick(Sender: TObject);
begin
LB2Nom.ItemIndex := LB2Libelle.ItemIndex;
end;

procedure TOF_PARAMOBLIG.LB2NomClick(Sender: TObject);
begin
LB2Libelle.ItemIndex := LB2Nom.ItemIndex;
end;

procedure TOF_PARAMOBLIG.LB2LibelleDblClick(Sender: TObject);
begin
bDeSelClick(Nil);
end;

procedure TOF_PARAMOBLIG.LB2NomDblClick(Sender: TObject);
begin
bDeSelClick(Nil);
end;

procedure TOF_PARAMOBLIG.bSelClick(Sender: TObject);
Var ChCol : TListeChps ;
    Ind : integer;
begin
if LB1Libelle.ItemIndex>=0 then
   BEGIN
   Ind:=LB1Libelle.ItemIndex ;
   ChCol:=TListeChps(LB1Libelle.Items.Objects[Ind]) ;
   LB2Libelle.Items.AddObject(ChCol.Intitule,ChCol) ;
   LB1Libelle.Items.Delete(Ind) ;
   if Ind<LB1Libelle.Items.Count then LB1Libelle.itemIndex:=Ind else LB1Libelle.itemIndex:=Ind-1 ;
   LB2Libelle.ItemIndex:=LB2Libelle.Items.Count-1 ;
   LB2Nom.Items.Add(ChCol.Champ);
   END ;
bModif := True;
bModifFin := True;
EnableButtons;
end;

procedure TOF_PARAMOBLIG.bDeSelClick(Sender: TObject);
Var ChCol : TListeChps ;
    Ind : integer;
    St : string ;
begin
if LB2Libelle.ItemIndex>=0 then
   BEGIN
   Ind:=LB2Libelle.ItemIndex ;
   ChCol:=TListeChps(LB2Libelle.Items.Objects[Ind]) ;
   if RG1.ItemIndex=0 then St:=ChCol.Intitule else St:=ChCol.Champ ;
   LB1Libelle.Items.AddObject(St,ChCol) ;
   LB2Libelle.Items.Delete(Ind) ;
   LB2Nom.Items.Delete(Ind);
   if Ind<LB2Libelle.Items.Count then LB2Libelle.itemIndex:=Ind else LB2Libelle.itemIndex:=Ind-1 ;
   END ;
bModif := True;
bModifFin := True;
EnableButtons;
end;

procedure TOF_PARAMOBLIG.bDeSelAllClick(Sender: TObject);
Var i_ind : integer ;
begin
for i_ind:=LB2Libelle.Items.Count-1 Downto 0 do
    begin
    LB2Libelle.ItemIndex:=i_ind;
    LB2Nom.ItemIndex:=i_ind;
    bDeSelClick(Sender);
    end;
end;

procedure TOF_PARAMOBLIG.bAnnuleClick(Sender: TObject);
begin
if bModif then
    if HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';Q;YNC;Y;C;','','')  = mrYes then
        bValideClick(nil);
Close;
end;

procedure TOF_PARAMOBLIG.bValideClick(Sender: TObject);
var ind1 : integer;
    TobTemp : TOB;
    sSel, sTmp1 : string;
begin
//if LB2Libelle.Items.Count = 0 then Exit;
for ind1 := 0 to LB2Libelle.Items.Count - 1 do
    begin
    TobTemp := TobFic.FindFirst(['GOB_NOMCHAMP'], [LB2Nom.Items.Strings[ind1]], False);
    sTmp1 := Copy(LB2Nom.Items.Strings[ind1], 0, Pos('_', LB2Nom.Items.Strings[ind1]) - 1);
    if TobTemp = nil then
        begin
        sSel := 'Select GOB_CODE from PARAMOBLIG Where GOB_CODE="' + CodePrec.Text + '" and ' +
                'GOB_NOMCHAMP="' + LB2Nom.Items.Strings[ind1] + '"';
        if ExisteSQL(sSel) then
            sSel := 'Update PARAMOBLIG Set GOB_OBLIGATOIRE="X" Where GOB_CODE="' + CodePrec.Text + '" and ' +
                    'GOB_NOMCHAMP="' + LB2Nom.Items.Strings[ind1] + '"'
            else
            sSel := 'Insert into PARAMOBLIG Values ("' + CodePrec.Text + '", ' +
                    '"' + LB2Nom.Items.Strings[ind1] + '", "' + sTmp1 + '", ' +
                    '"X", "", "", "", "", "", "", "")';
        end
        else
        sSel := 'Update PARAMOBLIG Set GOB_OBLIGATOIRE="X" Where GOB_CODE="' + CodePrec.Text + '" and ' +
                'GOB_NOMCHAMP="' + LB2Nom.Items.Strings[ind1] + '"';
    ExecuteSQL(sSel);
    end;
for ind1 := 0 to TobFic.Detail.Count - 1 do
    begin
    TobTemp := TobFic.Detail[ind1];
    if LB2Nom.Items.IndexOf(TobTemp.GetValue('GOB_NOMCHAMP')) < 0 then
        begin
        sSel := 'Update PARAMOBLIG Set GOB_OBLIGATOIRE="-" Where GOB_CODE="' + CodePrec.Text + '" and ' +
                'GOB_NOMCHAMP="' + TobTemp.GetValue('GOB_NOMCHAMP') + '"';
        ExecuteSQL(sSel);
        end;
    end;
end;

procedure TOF_PARAMOBLIG.RG1Click(Sender: TObject);
var i_ind,ind1 : integer ;
    ChCol : TListeChps ;
    St : String ;
begin
ind1:=LB1Libelle.ItemIndex ;
LB1Libelle.Sorted:=FALSE ;
For i_ind:=0 to LB1Libelle.Items.Count-1 do
    BEGIN
    ChCol:=TListeChps(LB1Libelle.Items.Objects[i_ind]) ;
    if RG1.ItemIndex=0 then St:=ChCol.Intitule else St:=ChCol.Champ ;
    LB1Libelle.Items[i_ind]:=St ;
    END ;
LB1Libelle.ItemIndex:=ind1 ;
LB1Libelle.Sorted:=TRUE ;
end;

procedure TOF_PARAMOBLIG.RG2Click(Sender: TObject);
begin
if RG2.ItemIndex = 0 then
    begin
    LB2Libelle.Visible := True;
    LB2Nom.Visible := False;
    end
    else
    begin
    LB2Libelle.Visible := False;
    LB2Nom.Visible := True;
    end;
end;

procedure TOF_PARAMOBLIG.EnableButtons;
begin
bSel.Enabled := (LB1Libelle.Items.Count > 0);
bSelAll.Enabled := (LB1Libelle.Items.Count > 0);
bDeSel.Enabled := (LB2Libelle.Items.Count > 0);
bDeSelAll.Enabled := (LB2Libelle.Items.Count > 0);
end;

//=======================================================================//

procedure TOF_PARAMOBLIG.LitDonnees;
var TOBCo : TOB ;
    sTmp1, sSel : string;
    Q : TQuery;
    ind1, iTableLigne : integer;
    ChCol,ChCol2 : TListeChps ;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//  Demande de validation
if bModif then
    if HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';Q;YNC;Y;C;','','')  = mrYes then
        bValideClick(nil);

CodePrec.Text := CodeActif.Text;
bModif := False;
for ind1:=0 to LB1Libelle.Items.Count-1 do
   TListeChps(LB1Libelle.Items.Objects[ind1]).Free ;
for ind1:=0 to LB2Libelle.Items.Count-1 do
   TListeChps(LB2Libelle.Items.Objects[ind1]).Free ;
LB1Libelle.Items.Clear;
LB1Nom.Items.Clear;
LB2Libelle.Items.Clear;
LB2Nom.Items.Clear;
LB1Libelle.Sorted:=False;
LB1Nom.Sorted:=False;
LB2Libelle.Sorted:=False;
LB2Nom.Sorted:=False;
TOBCo := TOB.Create('', nil, -1);
sSel := 'Select * from COMMUN where CO_TYPE="POB" and CO_CODE="' + CodeActif.Text + '"';
Q := OpenSQL(sSel, True,-1,'',true);
TOBCo.LoadDetailDB('COMMUN', '', '', Q, False);
Ferme(Q);
TableActive.Text := TOBCo.Detail[0].GetValue('CO_LIBRE');
TOBCo.Free;
TobFic.ClearDetail;

sSel := TableToPrefixe(Copy(TableActive.Text, 0, Pos(';', TableActive.Text) - 1));
Table := mcd.getTable(sSel);
FieldList := Table.Fields;
FieldList.Reset();
	While FieldList.MoveNext do
    begin
    if (FieldList.Current as IFieldCOM).Tipe = 'BOOLEAN' then Continue;
    if (FieldList.Current as IFieldCOM).Tipe = 'BLOB' then Continue;
    if (FieldList.Current as IFieldCOM).Obligatoire then Continue;

    ChCol:=TListeChps.Create ;
    if LB1Libelle.Items.Text = '' then
       begin
       LB1Nom.Items.Text     := (FieldList.Current as IFieldCOM).name;
       end
       else
       begin
       LB1Nom.Items.Add((FieldList.Current as IFieldCOM).name);
       end;
    LB1Libelle.Items.AddObject((FieldList.Current as IFieldCOM).Libelle,ChCol) ;
    ChCol.Champ:=(FieldList.Current as IFieldCOM).name;
    ChCol.Intitule:=(FieldList.Current as IFieldCOM).Libelle ;
    end;
EnableButtons;

sSel := 'Select * from PARAMOBLIG where GOB_CODE="' + CodeActif.Text + '" and GOB_OBLIGATOIRE="X"';
Q := OpenSQL(sSel, true,-1,'',true);
TobFic.LoadDetailDB('PARAMOBLIG', '', '', Q, False);
Ferme(Q);
if TobFic.Detail.Count>0 then
   begin
   for ind1 := 0 to TobFic.Detail.Count - 1 do
       begin
       sTmp1 := TobFic.Detail[ind1].GetValue('GOB_NOMCHAMP');
       if LB2Libelle.Items.Text = '' then
          begin
          LB2Nom.Items.Text     := sTmp1;
          end
          else
          begin
          LB2Nom.Items.Add(sTmp1);
          end;
       ChCol2:=TListeChps(LB1Libelle.Items.Objects[LB1Nom.Items.IndexOf(sTmp1)]) ;
       LB2Libelle.Items.AddObject(ChCol2.Intitule,ChCol2) ;
       LB1Libelle.Items.Delete(LB1Nom.Items.IndexOf(sTmp1));
       LB1Nom.Items.Delete(LB1Nom.Items.IndexOf(sTmp1));
       end;
   EnableButtons;
   end;
Rg1Click(Self);
Rg2Click(Self);
end;

Initialization
  registerclasses ( [ TOF_PARAMOBLIG ] ) ; 
end.
