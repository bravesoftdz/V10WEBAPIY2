{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PARAMCONFID ()
Mots clefs ... : TOF;PARAMCONFID
*****************************************************************}
Unit ParamConfid_Tof ;

Interface

Uses StdCtrls,
     ExtCtrls,
     Windows,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ELSE}
     MaineAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     Graphics,
     ImgList,
     HCtrls,
     HEnt1,
     HPanel,
     HMsgBox,
     HTB97,
     HZStream,
     UTOB,
     UTOF
{$IFDEF GRC}
     ,EntRT
{$ENDIF}
     ;

Type
  TOF_PARAMCONFID = Class (TOF)
    PanListe: THPanel;
    PanGroupe: THPanel;
    PanUser: THPanel;
    PanOptions: THPanel;
    PanRappel: THPanel;
    GRGroupe: THGrid;
    GRUser: THGrid;
    Img0 : TImage;
    Img1 : TImage;
    Img2 : TImage;
    Img3 : TImage;
    RB1 : TRadioButton;
    RB2 : TRadioButton;
    RB3 : TRadioButton;
    RB4 : TRadioButton;
    bAnnule: TToolbarButton97;
    bValide: TToolbarButton97;
    HValCB1: THValComboBox;
    TableActive: THEdit;
    CodeActif: THEdit;
    CodePrec: THEdit;
    Rappel: THLabel;
    TV : TTreeView;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    GRGroupeCell : string;
    GRUserCell : string;
    bModif : boolean;
    Memo : TMemo;
    ImgSel : TCustomImageList;
    TobTV : TOB;
    TobTVSauve : TOB;
    TobModif : TOB;
    TobSauve : TOB;
    TobConfid : TOB;
    TobGrp : TOB;
    TobUsr : TOB;
    TVCols : string;
    procedure FormResize(Sender: TObject);
    procedure HValCB1Click(Sender: TObject);
    procedure TVClick(Sender: TObject);
    procedure TVChange(Sender: TObject; Node: TTreeNode);
    procedure TVChanging(Sender: TObject; Node: TTreeNode;var AllowChange: Boolean);
    procedure TVMouseUp(Sender: TObject; Button: TMouseButton;
                        Shift: TShiftState; X, Y: Integer);
    procedure bAnnuleClick(Sender: TObject);
    procedure bValideClick(Sender: TObject);
    procedure GRGroupeMouseDown(Sender: TObject; Button: TMouseButton;
                                Shift: TShiftState; X, Y: Integer);
    procedure GRGroupeRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                               Chg: Boolean);
    procedure GRGroupeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                              Chg: Boolean);
    procedure GRGroupeCellEnter(Sender: TObject; var ACol, ARow: Integer;
                              var Cancel: Boolean);
    procedure GRGroupeCellExit(Sender: TObject; var ACol, ARow: Integer;
                             var Cancel: Boolean);
    procedure GRUserMouseDown(Sender: TObject; Button: TMouseButton;
                              Shift: TShiftState; X, Y: Integer);
    procedure GRUserCellEnter(Sender: TObject; var ACol, ARow: Integer;
                              var Cancel: Boolean);
    procedure GRUserCellExit(Sender: TObject; var ACol, ARow: Integer;
                             var Cancel: Boolean);
    procedure GRGrUsDblClick(Sender: TObject);
    procedure InitGrids;
    procedure GereIcones(Node : TTreeNode; Up : boolean);
    procedure LitDonnees;
    procedure LoadTV(var Ligne, Level : integer; TobMere : TOB);
    procedure ChargeGRPUSR(TobMere : TOB);
{$IFDEF GRC}
    procedure RTLoadTV(var Ligne, Level : integer; TobMere : TOB);
{$ENDIF}
    procedure RTAddTobTmp(var TobTmp:Tob; Nom,Libelle,TypeC : String);
    procedure RTDuplicClick(Sender: TObject);
  end ;


const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Voulez-vous enregistrer les modifications ?',
          {2}  'Vous devez saisir X ou un caractère blanc'
              );
ClientGC : String = 'G03';
ProspectGC : String = 'R01';
InfosClient : String = 'R02';
InfosProspect : String = 'R03';


procedure Entree_Confid;

Implementation
uses 
   CbpMCD
   ,CbpEnumerator
;

procedure Entree_Confid;
begin
V_PGI.ZoomOLE := true;  //Affichage en mode modal
AGLLanceFiche('GC','GCPARAMCONFID','','','');
V_PGI.ZoomOLE := false;
end;

procedure TOF_PARAMCONFID.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCONFID.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCONFID.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PARAMCONFID.OnLoad ;
var
    sSel : string;
    Q : TQuery;
begin
  Inherited ;
    Ecran.OnResize := FormResize;
    PanListe := THPanel(GetControl('PANLISTE'));
    PanGroupe := THPanel(GetControl('PANGROUPE'));
    PanUser := THPanel(GetControl('PANUSER'));
    PanOptions := THPanel(GetControl('PANOPTIONS'));
    PanRappel := THPanel(GetControl('PANRAPPEL'));
    Rappel := THLabel(GetControl('RAPPEL'));
//  PanOptions.Visible := False; // pour l'instant
    TV := TTreeView(GetControl('TV'));
    TV.OnClick := TVClick;
    TV.OnChange := TVChange;
    TV.OnChanging := TVChanging;
    TV.OnMouseUp := TVMouseUp;
    TV.ReadOnly := True;
    Img0 := TImage(GetControl('IMG0'));
    Img1 := TImage(GetControl('IMG1'));
    Img2 := TImage(GetControl('IMG2'));
    Img3 := TImage(GetControl('IMG3'));
    ImgSel := TCustomImageList.Create(TV);
    ImgSel.Add(Img0.Picture.Bitmap, nil);
    ImgSel.Add(Img1.Picture.Bitmap, nil);
    ImgSel.Add(Img2.Picture.Bitmap, nil);
    ImgSel.Add(Img3.Picture.Bitmap, nil);
    TV.Images := ImgSel;
//    LBLibelle.OnFlipSelection := LBLibelleFlipSelection;
    GRGroupe := THGrid(GetControl('GRGROUPE'));
    GRGroupe.OnRowEnter := GRGroupeRowEnter;
    GRGroupe.OnRowExit := GRGroupeRowExit;
    GRGroupe.OnMouseDown := GRGroupeMouseDown;
    GRGroupe.OnCellEnter := GRGroupeCellEnter;
    GRGroupe.OnCellExit := GRGroupeCellExit;
    GRGroupe.OnDblClick := GRGrUsDblClick;
    GRUser := THGrid(GetControl('GRUSER'));
    GRUser.OnMouseDown := GRUserMouseDown;
    GRUser.OnCellEnter := GRUserCellEnter;
    GRUser.OnCellExit := GRUserCellExit;
    GRUser.OnDblClick := GRGrUsDblClick;
    bAnnule := TToolbarButton97(GetControl('BFERME'));
    bAnnule.OnClick := bAnnuleClick;
    bAnnule.Visible := True;
    bValide := TToolbarButton97(GetControl('BVALIDER'));
    bValide.OnClick := bValideClick;
    RB1 := TRadioButton(GetControl('RB1'));
    RB2 := TRadioButton(GetControl('RB2'));
    RB3 := TRadioButton(GetControl('RB3'));
    RB4 := TRadioButton(GetControl('RB4'));
    HValCB1 := THValComboBox(GetControl('HVALCB1'));
    HValCB1.OnClick := HValCB1Click;
    TableActive := THEdit(GetControl('TABLEACTIVE'));
    CodeActif := THEdit(GetControl('CODEACTIF'));
    CodePrec := THEdit(GetControl('CODEPREC'));
    Memo := TMemo(GetControl('MEMO'));
    Memo.Align := alClient;
    TVCols := 'Libelle;Libelle;Libelle;Libelle;Libelle;Libelle;Libelle;Libelle;Libelle;';
    TableActive.Text := '';
    CodeActif.Text := '';
    CodePrec.Text := '';
    bModif := False;
//  Retouche de l'affichage
    PanListe.Align := alLeft;
    PanListe.Width := Trunc(Ecran.ClientWidth / 3);
    PanGroupe.Align := alLeft;
    PanGroupe.Width := Trunc((Ecran.ClientWidth - PanListe.Width) / 2);
    PanUser.Align := alClient;
    GRGroupe.ColAligns[1] := taCenter;
    GRGroupe.ColAligns[2] := taCenter;
    GRGroupe.ColLengths[0] := -1;
    GRGroupe.ColWidths[1] := Trunc((GRGroupe.Width - 20) * 0.20);
    GRGroupe.ColWidths[2] := Trunc((GRGroupe.Width - 20) * 0.20);
    GRGroupe.ColWidths[0] := Trunc((GRGroupe.Width - 20) * 0.60);
    if GRGroupe.ColWidths[1] < 60 then
        begin
        GRGroupe.ColWidths[1] := 60;
        GRGroupe.ColWidths[2] := 60;
        GRGroupe.ColWidths[0] := GRGroupe.ClientWidth - 120 - GetSystemMetrics(SM_CYVSCROLL);
        end;
    GRUser.ColAligns[1] := taCenter;
    GRUser.ColAligns[2] := taCenter;
    GRUser.ColLengths[0] := -1;
    GRUser.ColWidths[0] := Trunc((GRUser.Width - 20) * 0.60);
    GRUser.ColWidths[1] := Trunc((GRUser.Width - 20) * 0.20);
    GRUser.ColWidths[2] := Trunc((GRUser.Width - 20) * 0.20);
    if GRUser.ColWidths[1] < 60 then
        begin
        GRUser.ColWidths[1] := 60;
        GRUser.ColWidths[2] := 60;
        GRUser.ColWidths[0] := GRUser.ClientWidth - 120 - GetSystemMetrics(SM_CYVSCROLL);
        end;
    TobTV := TOB.Create('', nil, -1);
    TobTVSauve := TOB.Create('', nil, -1);
    TobModif := TOB.Create('', nil, -1);
    TobSauve := TOB.Create('', nil, -1);
    //TobConfid := TOB.Create('', nil, -1);
//  Chargement des tables groupes et utilisateurs
    TobGrp := TOB.Create('', nil, -1);
    sSel := 'Select * from USERGRP order by UG_NUMERO';
    Q := OpenSQL(sSel, True,-1, '', True);
    TobGrp.LoadDetailDB('USERGRP', '', '', Q, False);
    Ferme(Q);
    TobGrp.Detail[0].AddChampSup('GRPVISIBLE', True);
    TobGrp.Detail[0].AddChampSup('GRPENABLE', True);

    TobUsr := TOB.Create('', nil, -1);
    sSel := 'Select * from UTILISAT';
    Q := OpenSQL(sSel, True,-1, '', True);
    TobUsr.LoadDetailDB('UTILISAT', '', '', Q, False);
    Ferme(Q);
    TobUsr.Detail[0].AddChampSup('USRVISIBLE', True);
    TobUsr.Detail[0].AddChampSup('USRENABLE', True);
end ;

procedure TOF_PARAMCONFID.OnArgument (S : String ) ;
var      Combo : THValComboBox;
         Plus : String;
begin
  Inherited ;
  // mcd 22/05/02 pour prendre en compte les domaines
  Combo:=THValComboBox(GetControl('HVALCB1'));
  
  //gm 06/10/03
  if (ctxaffaire in V_PGI.PGIContexte) then
    Plus:='AND (CO_CODE like "A%" OR CO_CODE like "B%"'
  else
    if (ctxgcaff in V_PGI.PGIContexte) then
      Plus:='AND (CO_CODE like "G%" OR CO_CODE like "A%"'
    else
      Plus:='AND (CO_CODE like "G%"';

  if ctxGRC in V_PGI.PGIContexte then Plus :=Plus + ' OR CO_CODE like "R%")'
  else  Plus :=Plus + ')';
  Plus:=Plus+' AND (CO_LIBRE Not Like "%&C&%")';
  {$IFDEF CHR}
  plus:='AND (CO_CODE like "G%" OR CO_CODE like "H%" OR CO_CODE like "R%") AND (CO_LIBRE Not Like "%&C&%")';
  {$ENDIF}
  Combo.plus:=Plus;
  Combo.ExHaustif:=Explus;
end ;

procedure TOF_PARAMCONFID.OnClose ;
begin
  Inherited ;
TobTV.Free;
TobTVSauve.Free;
TobModif.Free;
TobSauve.Free;
TobGrp.Free;
TobUsr.Free;
end ;

procedure TOF_PARAMCONFID.FormResize(Sender: TObject);
begin
PanListe.Width := Trunc(Ecran.ClientWidth / 3);
PanGroupe.Width := Trunc((Ecran.ClientWidth - PanListe.Width) / 2);
GRGroupe.ColWidths[1] := Trunc((GRGroupe.Width - 20) * 0.20);
GRGroupe.ColWidths[2] := Trunc((GRGroupe.Width - 20) * 0.20);
GRGroupe.ColWidths[0] := Trunc((GRGroupe.Width - 20) * 0.60);
if GRGroupe.ColWidths[1] < 60 then
    begin
    GRGroupe.ColWidths[1] := 60;
    GRGroupe.ColWidths[2] := 60;
    GRGroupe.ColWidths[0] := GRGroupe.ClientWidth - 120 - GetSystemMetrics(SM_CYVSCROLL);
    end;
GRUser.ColWidths[0] := Trunc((GRUser.Width - 20) * 0.60);
GRUser.ColWidths[1] := Trunc((GRUser.Width - 20) * 0.20);
GRUser.ColWidths[2] := Trunc((GRUser.Width - 20) * 0.20);
if GRUser.ColWidths[1] < 60 then
    begin
    GRUser.ColWidths[1] := 60;
    GRUser.ColWidths[2] := 60;
    GRUser.ColWidths[0] := GRUser.ClientWidth - 120 - GetSystemMetrics(SM_CYVSCROLL);
    end;
end;

procedure TOF_PARAMCONFID.HValCB1Click(Sender: TObject);
var Rep : integer;
    bDuplic : TToolbarButton97;
begin
//  Demande de validation
if bModif then
    begin
    Rep := HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';Q;YNC;Y;C;','','');
    Case Rep of
        mrYes    : bValideClick(nil);
        mrCancel : begin
                   HValCB1.ItemIndex := HValCB1.Values.IndexOf(CodeActif.Text);
                   Exit;
                   end;
        end;
    end;
bModif := False;
CodeActif.Text := HValCB1.Value;
if CodePrec.Text = '' then CodePrec.Text := CodeActif.Text;
GRGroupe.OnRowEnter := nil;
GRGroupe.OnRowExit := nil;
LitDonnees;
GRGroupe.OnRowEnter := GRGroupeRowEnter;
GRGroupe.OnRowExit := GRGroupeRowExit;
if ctxGRC in V_PGI.PGIContexte then
    begin
    // bouton Duplication pour clients ou prospects
    bDuplic := TToolbarButton97(GetControl('BDUPLIC'));
    if (CodeActif.Text = ClientGC) or (CodeActif.Text = InfosClient) or
       (CodeActif.Text = ProspectGC) or (CodeActif.Text = InfosProspect) then
        begin
        bDuplic.Visible:=True;
        if (CodeActif.Text = ClientGC) or (CodeActif.Text = InfosClient) then
            bDuplic.Hint:='Duplication de la confidentialité des Prospects'
        else
            bDuplic.Hint:='Duplication de la confidentialité des Clients';
        bDuplic.OnClick:=RTDuplicClick;
        end
    else
        bDuplic.Visible:=False;
    end;
end;

procedure TOF_PARAMCONFID.TVClick(Sender: TObject);
begin
;
end;

procedure TOF_PARAMCONFID.TVChange(Sender: TObject; Node: TTreeNode);
begin
if Node.ImageIndex = 3 then Exit;
Node.SelectedIndex := Node.ImageIndex;
Rappel.Caption := 'Paramétrage pour : ' + Node.Text;
Rappel.Left := PanGroupe.Left + Trunc((PanGroupe.Width + PanUser.Width - Rappel.Width) / 2);
GRGroupe.Col := 1;
GRUser.Col := 1;
TobConfid := TobModif.FindFirst(['GOB_NOMCHAMP'], [TOB(Node.Data).GetValue('Nom')], False);
if TobConfid = nil then
    begin
    TOB(Node.Data).PutValue('Selected', False);
    GRGroupe.VidePile(True);
    GRUser.VidePile(True);
    GRGroupe.Enabled := False;
    GRUser.Enabled := False;
    Exit;
    end;
InitGrids;
end;

procedure TOF_PARAMCONFID.TVChanging(Sender: TObject; Node: TTreeNode; var AllowChange: Boolean);
var
   TobTmpGrp, TobTmpUsr : TOB;
   ind1, ind2 : integer;
begin
if TobConfid = nil then Exit;
for ind1 := 1 to GRGroupe.RowCount - 1 do
    begin
    TobTmpGrp := TobConfid.FindFirst(['UG_LIBELLE'], [GRGroupe.Cells[0, ind1]], False);
    if TobTmpGrp <> nil then
        begin
        TobTmpGrp.PutValue('GRPVISIBLE', GRGroupe.Cells[1, ind1]);
        TobTmpGrp.PutValue('GRPENABLE', GRGroupe.Cells[2, ind1]);
        for ind2 := 1 to GRUser.RowCount - 1 do
            begin
            TobTmpUsr := TobTmpGrp.FindFirst(['US_LIBELLE'], [GRUser.Cells[0, ind2]], True);
            if TobTmpUsr <> nil then
                begin
                TobTmpUsr.PutValue('USRVISIBLE', GRUser.Cells[1, ind2]);
                TobTmpUsr.PutValue('USRENABLE', GRUser.Cells[2, ind2]);
                end;
            end;
        end;
    end;
end;

procedure TOF_PARAMCONFID.TVMouseUp(Sender: TObject; Button: TMouseButton;
                                    Shift: TShiftState; X, Y: Integer);
var
   TobTmpOblig, TobTmpSauve : TOB;
   Node : TTreeNode;
   ind1 : integer;
begin
Node := TV.GetNodeAt(X, Y);
if Node = Nil then Exit;
if Node.ImageIndex = 3 then Exit;
if not TOB(Node.Data).GetValue('Modif') then
    begin
    Node.ImageIndex := 3;
    Exit;
    end;
if Button <> mbRight then Exit;
Node.Selected := True;
TobTmpOblig := TobModif.FindFirst(['GOB_NOMCHAMP'], [TOB(Node.Data).GetValue('Nom')], False);
if TobTmpOblig = nil then
    begin
    TOB(Node.Data).PutValue('Selected', True);
    if TOB(Node.Data).Detail.Count <> 0 then
        begin
        TOB(Node.Data).ClearDetail;
        Node.DeleteChildren;
        end;
    TobTmpOblig := TOB.Create('PARAMOBLIG', TobModif, -1);
    TobTmpOblig.InitValeurs;
    TobTmpOblig.PutValue('GOB_CODE', CodeActif.Text);
    TobTmpOblig.PutValue('GOB_NOMCHAMP', TOB(Node.Data).GetValue('Nom'));
    TobTmpOblig.PutValue('GOB_PREFIXE',
                         Copy(TOB(Node.Data).GetValue('Nom'),
                              0, Pos('_', TOB(Node.Data).GetValue('Nom')) - 1));
    TobTmpOblig.AddChampSup('GOB_USVISIBLE', True);
    TobTmpOblig.AddChampSup('GOB_USENABLE', True);
    ChargeGRPUSR(TobTmpOblig);
    TobConfid := TobModif.FindFirst(['GOB_NOMCHAMP'], [TOB(Node.Data).GetValue('Nom')], False);
    InitGrids;
    end
    else
    begin
    TobTmpOblig.Free;
    TOB(Node.Data).PutValue('Selected', False);
    TobTmpSauve := TobTVSauve.FindFirst(['Libelle'], [Node.Text], True);
    if TobTmpSauve <> nil then
        if TobTmpSauve.Detail.Count <> 0 then
            begin
            TOB(Node.Data).Dupliquer(TobTmpSauve, True, True);
            for ind1 := 0 to TOB(Node.Data).Detail.Count - 1 do
                TOB(Node.Data).Detail[ind1].PutTreeView(TV, Node, TVCols);
            GereIcones(Node, False);
            end;
    GRGroupe.VidePile(True);
    GRUser.VidePile(True);
    GRGroupe.Enabled := False;
    GRUser.Enabled := False;
    end;
TobConfid := TobModif.FindFirst(['GOB_NOMCHAMP'], [TOB(Node.Data).GetValue('Nom')], False);
if TOB(Node.Data).GetValue('Selected') then
    Node.ImageIndex := 1
    else
    Node.ImageIndex := 0;
GereIcones(Node, True);
Node.SelectedIndex := Node.ImageIndex;
end;

procedure TOF_PARAMCONFID.bAnnuleClick(Sender: TObject);
var Rep : integer;
begin
if bModif then
    begin
    Rep := HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';Q;YNC;Y;C;','','');
    Case Rep of
        mrYes : begin
                bValideClick(nil);
                Close;
                end;
        mrNo  : Close;
        end;
    end;
end;

procedure TOF_PARAMCONFID.bValideClick(Sender: TObject);
var Sql : string;
    Cancel : boolean;

    procedure EnregistreData(TN : TTreeNode);
    var
        ind1, ind2, ind3 : integer;
        GrVisible, GrEnable, UsVisible, UsEnable : string;
        TobTemp, TobTmpGrp, TobTmpUsr : TOB;
        TNTemp : TTreeNode;
    begin
    for ind1 := 0 to TN.Count - 1 do
        begin
        TNTemp := TN.Item[ind1];
        if TNTemp.Count <> 0 then
            EnregistreData(TNTemp)
            else
            begin
            if not TOB(TNTemp.Data).GetValue('Selected') then Continue;
            TobTemp := TobModif.FindFirst(['GOB_NOMCHAMP'], [TOB(TNTemp.Data).GetValue('Nom')], False);
            GrVisible := '';
            UsVisible := '';
            GrEnable := '';
            UsEnable := '';
            for ind2 := 0 to TobTemp.Detail.Count - 1 do
                begin
                TobTmpGrp := TobTemp.Detail[ind2];
                if TobTmpGrp.GetValue('GRPVISIBLE') = 'X' then
                    GrVisible := GrVisible + TobTmpGrp.GetValue('UG_GROUPE') + ';';
                if TobTmpGrp.GetValue('GRPENABLE') = 'X' then
                    GrEnable := GrEnable + TobTmpGrp.GetValue('UG_GROUPE') + ';';
                for ind3 := 0 to TobTmpGrp.Detail.Count - 1 do
                    begin
                    TobTmpUsr := TobTmpGrp.Detail[ind3];
                    if TobTmpUsr.GetValue('USRVISIBLE') = 'X' then
                        UsVisible := UsVisible + TobTmpUsr.GetValue('US_UTILISATEUR') + ';';
                    if TobTmpUsr.GetValue('USRENABLE') = 'X' then
                        UsEnable := UsEnable + TobTmpUsr.GetValue('US_UTILISATEUR') + ';';
                    end;
                end;
            TobTemp.PutValue('GOB_GRVISIBLE', GrVisible);
            TobTemp.PutValue('GOB_GRENABLE', GrEnable);
            TobTemp.PutValue('GOB_USVISIBLE', UsVisible);
            TobTemp.PutValue('GOB_USENABLE', UsEnable);
            if Length(UsVisible) > 200 then
                begin
                TobTemp.PutValue('GOB_USVISIBLE1', Copy(UsVisible, 1, 200));
                TobTemp.PutValue('GOB_USVISIBLE2', Copy(UsVisible, 201, Length(UsVisible) - 200));
                end
                else
                begin
                TobTemp.PutValue('GOB_USVISIBLE1', UsVisible);
                TobTemp.PutValue('GOB_USVISIBLE2', '');
                end;
            if Length(UsEnable) > 200 then
                begin
                TobTemp.PutValue('GOB_USENABLE1', Copy(UsEnable, 1, 200));
                TobTemp.PutValue('GOB_USENABLE2', Copy(UsEnable, 201, Length(UsEnable) - 200));
                end
                else
                begin
                TobTemp.PutValue('GOB_USENABLE1', UsEnable);
                TobTemp.PutValue('GOB_USENABLE2', '');
                end;
            if RB1.Checked then TobTemp.PutValue('GOB_PRIOCONFID', '1') else
              if RB2.Checked then TobTemp.PutValue('GOB_PRIOCONFID', '2') else
                if RB3.Checked then TobTemp.PutValue('GOB_PRIOCONFID', '3') else
                  if RB4.Checked then TobTemp.PutValue('GOB_PRIOCONFID', '4') ;
            TobTemp.SetAllModifie(True);
            TobTemp.InsertOrUpdateDB;
            end;
        end;
    end;
begin
GRGroupeRowExit(nil, GRGroupe.Row, Cancel, False);
TVChanging(nil, TV.Selected, Cancel);

Sql := 'Delete from PARAMOBLIG where GOB_CODE="' + CodeActif.Text + '" and GOB_OBLIGATOIRE<>"X"';
ExecuteSql(Sql);

EnregistreData(TV.Items[0]);

bModif := False;
end;

procedure TOF_PARAMCONFID.GRGroupeMouseDown(Sender: TObject; Button: TMouseButton;
                                            Shift: TShiftState; X, Y: Integer);
var
   ACol, ARow : integer;
begin
GRGroupe.MouseToCell(X, Y, ACol, ARow);
if ACol = 0 then
    begin
    GRGroupe.Row := ARow;
    GRGroupe.Col := 1;
    end;
end;

procedure TOF_PARAMCONFID.GRGroupeRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var
   TobTmpGrp, TobTmpUsr : TOB;
   ind1 : integer;
begin
if GRGroupe.Cells[0, GRGroupe.Row] = '' then Exit;
TobTmpGrp := TobConfid.FindFirst(['UG_LIBELLE'], [GRGroupe.Cells[0, Ou]], False);
if TobTmpGrp = nil then Exit;
//GRUser.VidePile(False);
for ind1 := 1 to GRUser.RowCount - 1 do
    GRUser.Rows[ind1].Clear;
if TobTmpGrp.Detail.Count <> 0 then
    begin
    GRUser.RowCount := TobTmpGrp.Detail.Count + 1;
    GRUser.Enabled := True;
    end
    else
    begin
    GRUser.RowCount := 2;
    GRUser.Enabled := False;
    end;

for ind1 := 0 to TobTmpGrp.Detail.Count - 1 do
    begin
    TobTmpUsr := TobTmpGrp.Detail[ind1];
    GRUser.Cells[0, ind1 + 1] := TobTmpUsr.GetValue('US_LIBELLE');
    GRUser.Cells[1, ind1 + 1] := TobTmpUsr.GetValue('USRVISIBLE');
    GRUser.Cells[2, ind1 + 1] := TobTmpUsr.GetValue('USRENABLE');
    end;
GRUserCell := GRUser.Cells[GRUser.Col, GRUser.Row];
end;

procedure TOF_PARAMCONFID.GRGroupeRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var
   TobTmpGrp, TobTmpUsr : TOB;
   ind1 : integer;
begin
if TobConfid=nil then Exit;
TobTmpGrp := TobConfid.FindFirst(['UG_LIBELLE'], [GRGroupe.Cells[0, Ou]], False);
if TobTmpGrp <> nil then
    begin
    for ind1 := 1 to GRUser.RowCount - 1 do
        begin
        TobTmpUsr := TobTmpGrp.FindFirst(['US_LIBELLE'], [GRUser.Cells[0, ind1]], True);
        if TobTmpUsr <> nil then
            begin
            TobTmpUsr.PutValue('USRVISIBLE', GRUser.Cells[1, ind1]);
            TobTmpUsr.PutValue('USRENABLE', GRUser.Cells[2, ind1]);
            end;
        end;
    end;
end;

procedure TOF_PARAMCONFID.GRGroupeCellEnter(Sender: TObject; var ACol, ARow: Integer;
                                            var Cancel: Boolean);
begin
GRGroupeCell := GRGroupe.Cells[GRGroupe.Col, GRGroupe.Row];
end;

procedure TOF_PARAMCONFID.GRGroupeCellExit(Sender: TObject; var ACol, ARow: Integer;
                                           var Cancel: Boolean);
begin
GRGroupe.Cells[ACol, ARow] := UpperCase(GRGroupe.Cells[ACol, ARow]);
if (GRGroupe.Cells[ACol, ARow] <> 'X') and (GRGroupe.Cells[ACol, ARow] <> ' ') then
    begin
    HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[2]+';W;O;O;O;','','');
    Cancel := True;
    Exit;
    end;
if GRGroupeCell <> GRGroupe.Cells[ACol, ARow] then bModif := True;
end;

procedure TOF_PARAMCONFID.GRUserMouseDown(Sender: TObject; Button: TMouseButton;
                                          Shift: TShiftState; X, Y: Integer);
var
   ACol, ARow : integer;
begin
GRUser.MouseToCell(X, Y, ACol, ARow);
if ACol = 0 then
    begin
    GRUser.Row := ARow;
    GRUser.Col := 1;
    end;
end;

procedure TOF_PARAMCONFID.GRUserCellEnter(Sender: TObject; var ACol, ARow: Integer;
                                          var Cancel: Boolean);
begin
GRUserCell := GRUser.Cells[GRUser.Col, GRUser.Row];
end;

procedure TOF_PARAMCONFID.GRUserCellExit(Sender: TObject; var ACol, ARow: Integer;
                                         var Cancel: Boolean);
begin
GRUser.Cells[ACol, ARow] := UpperCase(GRUser.Cells[ACol, ARow]);
if (GRUser.Cells[ACol, ARow] <> 'X') and (GRUser.Cells[ACol, ARow] <> ' ') then
    begin
    HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[2]+';W;O;O;O;','','');
    Cancel := True;
    Exit;
    end;
if GRUserCell <> GRUser.Cells[ACol, ARow] then bModif := True;
end;

procedure TOF_PARAMCONFID.GRGrUsDblClick(Sender: TObject);
begin
if THGrid(Sender).Cells[THGrid(Sender).Col, THGrid(Sender).Row] = ' ' then
   THGrid(Sender).Cells[THGrid(Sender).Col, THGrid(Sender).Row] := 'X'
   else
   THGrid(Sender).Cells[THGrid(Sender).Col, THGrid(Sender).Row] := ' ';
end;

procedure TOF_PARAMCONFID.InitGrids;
var
    ind1 : integer;
    TobTmpGrp : TOB;
    Cancel : boolean;

begin
GRGroupe.Enabled := True;
GRUser.Enabled := True;
GRGroupe.VidePile(True);
GRGroupe.RowCount := TobConfid.Detail.Count + 1;
for ind1 := 0 to TobConfid.Detail.Count - 1 do
    begin
    TobTmpGrp := TobConfid.Detail[ind1];
    GRGroupe.Cells[0, ind1 + 1] := TobTmpGrp.GetValue('UG_LIBELLE');
    GRGroupe.Cells[1, ind1 + 1] := TobTmpGrp.GetValue('GRPVISIBLE');
    GRGroupe.Cells[2, ind1 + 1] := TobTmpGrp.GetValue('GRPENABLE');
    end;
GRGroupe.Row := 1;
GRGroupeRowEnter(nil, 1, Cancel, False);
GRGroupeCell := GRGroupe.Cells[GRGroupe.Col, GRGroupe.Row];
end;

procedure TOF_PARAMCONFID.GereIcones(Node : TTreeNode; Up : boolean);
var
   ind1, NB0, NB1 : integer;
begin
if Up then
    begin
    if not TOB(Node.Data).GetValue('Modif') then
        begin
        Node.ImageIndex := 3;
        Exit;
        end;
    if Node.Count <> 0 then
        begin
        NB0 := 0;
        NB1 := 0;
        for ind1 := 0 to Node.Count - 1 do
            if Node.Item[ind1].ImageIndex = 0 then
                Inc(NB0)
                else
                Inc(NB1);
        Node.ImageIndex := 0;
        if (NB0 <> 0) and (NB1 <> 0) then
            Node.ImageIndex := 2;
        end;
    if Node.Parent <> nil then GereIcones(Node.Parent, True);
    end
    else
    begin
    for ind1 := 0 to Node.Count - 1 do
        begin
        if TOB(Node.Item[ind1].Data).GetValue('Selected') then
            Node.Item[ind1].ImageIndex := 1
            else
            Node.Item[ind1].ImageIndex := 0;
        if not TOB(Node.Item[ind1].Data).GetValue('Modif') then
            Node.Item[ind1].ImageIndex := 3;
        Node.Item[ind1].SelectedIndex := Node.Item[ind1].ImageIndex;
        GereIcones(Node.Item[ind1], False);
        end;
//    if Node.Parent <> nil then GereIcones(Node.Parent, True);
    end;
end;

procedure TOF_PARAMCONFID.LitDonnees;
var
    sSel : string;
    Q : TQuery;
    ind1, Level, Ligne : integer;
    TobTemp : TOB;
{$IFDEF EAGLCLIENT}
    Dfm : string;
{$ELSE}
    InputStream : TMemoryStream ;
{$ENDIF}

begin
CodePrec.Text := CodeActif.Text;
bModif := False;
TV.TopItem := nil;
TobTV.ClearDetail;
TobTVSauve.ClearDetail;
TobModif.ClearDetail;
if TobConfid <> nil then
    begin
    TobConfid.ClearDetail;
    TobConfid := nil;
    end;
GRGroupe.VidePile(True);
GRUser.VidePile(True);
TobTemp := Tob.Create('', nil, -1);
sSel := 'Select * from COMMUN where CO_TYPE="POB" and CO_CODE="' + CodeActif.Text + '"';
Q := OpenSQL(sSel, True,-1, '', True);
TobTemp.LoadDetailDB('COMMUN', '', '', Q, False);
Ferme(Q);
sSel := TobTemp.Detail[0].GetValue('CO_LIBRE');
ReadTokenSt(sSel);
TableActive.Text := ReadTokenSt(sSel);
TobTemp.Free;
RB1.Checked := (sSel = '1');
RB2.Checked := (sSel = '2');
RB3.Checked := (sSel = '3');
RB4.Checked := (sSel = '4');

sSel := 'Select GOB_CODE, GOB_NOMCHAMP, GOB_PREFIXE, GOB_GRVISIBLE, GOB_GRENABLE, GOB_USVISIBLE1, ' +
        'GOB_USVISIBLE2, GOB_USENABLE1, GOB_USENABLE2,GOB_PRIOCONFID from PARAMOBLIG where GOB_CODE="' + CodeActif.Text + '" and' +
        ' (GOB_GRVISIBLE<>"" or GOB_GRENABLE<>"" or GOB_USVISIBLE1<>"" or GOB_USENABLE1<>"")';
Q := OpenSQL(sSel, true,-1,'',true);
TobModif.LoadDetailDB('PARAMOBLIG', '', '', Q, False);
Ferme(Q);

for ind1 := 0 to TobModif.Detail.Count - 1 do
    begin
    TobTemp := TobModif.Detail[ind1];
    if ind1=0 then
       begin
       RB1.Checked := (TobTemp.GetValue('GOB_PRIOCONFID')='1');
       RB2.Checked := (TobTemp.GetValue('GOB_PRIOCONFID')='2');
       RB3.Checked := (TobTemp.GetValue('GOB_PRIOCONFID')='3');
       RB4.Checked := (TobTemp.GetValue('GOB_PRIOCONFID')='4');
       end;
    TobTemp.AddChampSup('GOB_USVISIBLE', True);
    TobTemp.AddChampSup('GOB_USENABLE', True);
    TobTemp.PutValue('GOB_USVISIBLE', Trim( TobTemp.GetValue('GOB_USVISIBLE1') +
                                            TobTemp.GetValue('GOB_USVISIBLE2')));
    TobTemp.PutValue('GOB_USENABLE', Trim( TobTemp.GetValue('GOB_USENABLE1') +
                                           TobTemp.GetValue('GOB_USENABLE2')));
    ChargeGRPUSR(TobTemp);
    end;
TobSauve.Dupliquer(TobModif, True, True);
if (CodeActif.Text <> InfosClient ) and (CodeActif.Text <> InfosProspect ) then
    begin
    sSel := 'Select * from FORMES where DFM_FORME="' + TableActive.Text + '"';
    Q := OpenSql(sSel, True,-1,'',true);
    {$IFDEF EAGLCLIENT}
    Dfm := Q.FindField('DFM_DATA').AsString ;
    AGLDecompressString(Dfm);
    sSel:=Dfm;
    {$ELSE}
    InputStream := TMemoryStream.Create ; InputStream.Clear;
    TBlobField(Q.FindField('DFM_DATA')).SaveToStream(InputStream) ;
    DecompressTextObjectStream(InputStream);
    SetString(sSel, PChar(InputStream.Memory), InputStream.Size);
    InputStream.Free;
    {$ENDIF}
    Ferme(Q);
    Memo.Lines.Text := sSel;
    sSel := Memo.Lines.Strings[0];
    System.Delete(sSel, 1, Pos(' ', sSel));
    System.Delete(sSel, Pos(':', sSel), 255);
    TobTV.AddChampSup('Nom', True);
    TobTV.AddChampSup('Type', True);
    TobTV.AddChampSup('Libelle', True);
    TobTV.AddChampSup('Modif', True);
    TobTV.AddChampSup('Selected', True);
    TobTV.PutValue('Nom', sSel);
    sSel := Memo.Lines.Strings[0];
    System.Delete(sSel, 1, Pos(':', sSel));
    TobTV.PutValue('Type', Trim(sSel));
    TobTV.PutValue('Libelle', 'Fiche');
    TobTV.PutValue('Modif', False);
    TobTV.PutValue('Selected', False);
    Ligne := 1;
    Level := 1;
    LoadTV(Ligne, Level, TobTV);
    end
{$IFDEF GRC}
else
    begin
    TobTV.AddChampSup('Nom', True);
    TobTV.AddChampSup('Type', True);
    TobTV.AddChampSup('Libelle', True);
    TobTV.AddChampSup('Modif', True);
    TobTV.AddChampSup('Selected', True);
    TobTV.PutValue('Nom', 'RTPARAMCL');
    TobTV.PutValue('Type', 'TFVierge');
    TobTV.PutValue('Libelle', 'Fiche');
    TobTV.PutValue('Modif', False);
    TobTV.PutValue('Selected', False);
    Ligne := 1;
    Level := 1;
    RTLoadTV(Ligne, Level, TobTV);
    end
{$ENDIF}
    ;
TobTVSauve.Dupliquer(TobTV, True, True);
TobTV.PutTreeView(TV, nil, TVCols);
for ind1 := TV.Items.Count - 1 downto 0 do
    if TOB(TV.Items.Item[ind1].Data).GetValue('Selected') then
        if TV.Items.Item[ind1].Count <> 0 then
            begin
            TOB(TV.Items.Item[ind1].Data).ClearDetail;
            TV.Items.Item[ind1].DeleteChildren;
            end;
for ind1 := 0 to TV.Items.Count - 1 do
    begin
    if not TOB(TV.Items.Item[ind1].Data).GetValue('Modif') then
        begin
        TV.Items.Item[ind1].ImageIndex := 3;
        TV.Items.Item[ind1].SelectedIndex := 3;
        end
    else if TOB(TV.Items.Item[ind1].Data).GetValue('Selected') then
        TV.Items.Item[ind1].ImageIndex := 1
        else
        TV.Items.Item[ind1].ImageIndex := 0;
    if TV.Items.Item[ind1].Parent <> nil then GereIcones(TV.Items.Item[ind1], True);
    end;
TV.TopItem.Expand(False);
end;

procedure TOF_PARAMCONFID.LoadTV(var Ligne, Level : integer; TobMere : TOB);
var
   sTmp1, sTmp2, sSave : string;
   ind1, iTableLigne, iPos : integer;
   TobTmp : TOB;
   Trouve : boolean;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
while (Ligne <= Memo.lines.Count - 1) do
    begin
    sTmp1 := Trim(Memo.Lines.Strings[Ligne]);
    if (Pos('inherited', sTmp1) <> 0) or
       (Pos('object', sTmp1) <> 0) then
        begin
        sSave := sTmp1;
        System.Delete(sTmp1, 1, Pos(' ', sTmp1));
        System.Delete(sTmp1, Pos(':', sTmp1), 255);
        TobTmp := TOB.Create('', TobMere, -1);
        TobTmp.AddChampSup('Nom', True);
        TobTmp.AddChampSup('Type', True);
        TobTmp.AddChampSup('Libelle', True);
        TobTmp.AddChampSup('Modif', True);
        TobTmp.AddChampSup('Selected', True);
        TobTmp.PutValue('Nom', sTmp1);
        TobTmp.PutValue('Libelle', sTmp1);
        sTmp1 := Memo.Lines.Strings[Ligne];
        System.Delete(sTmp1, 1, Pos(':', sTmp1));
        if Pos('[', sTmp1) <> 0 then System.Delete(sTmp1, Pos('[', sTmp1), 255);
        TobTmp.PutValue('Type', Trim(sTmp1));
        TobTmp.PutValue('Selected', False);
        Inc(Ligne);
        Inc(Level);
        LoadTV(Ligne, Level, TobTmp);
        if TobModif.FindFirst(['GOB_NOMCHAMP'], [TobTmp.GetValue('Nom')], False) <> nil then
            TobTmp.PutValue('Selected', True);

        TobTmp.PutValue('Modif', True);
        sTmp1 := TobTmp.GetValue('Nom');
        if Copy(sTmp1, 0, 6) = 'Dock97' then TobTmp.PutValue('Libelle', 'Obj. Système');
        if Pos('inherited', sSave) <> 0 then
            begin
            TobTmp.PutValue('Modif', False);
            if TobTmp.Detail.Count = 0 then TobTmp.Free;
            end
        else if (Pos('object', sSave) <> 0) then
            begin
            if TobTmp.GetValue('Type') = 'TPopupMenu' then
                begin
                Trouve := False;
                for ind1 := 0 to Memo.Lines.Count - 1 do
                    if Pos('DropdownMenu = ' + TobTmp.GetValue('Nom'), Memo.Lines.Strings[ind1]) <> 0 then
                        begin
                        Trouve := True;
                        Break;
                        end;
                if Trouve then
                    begin
                    while Pos('object', Memo.Lines.Strings[ind1]) = 0 do
                        Dec(ind1);
                    while Pos('end', Memo.Lines.Strings[ind1]) = 0 do
                        begin
                        if (Pos(' Caption', Memo.Lines.Strings[ind1]) <> 0) or
                           (Pos(' Hint', Memo.Lines.Strings[ind1]) <> 0) then
                            begin
                            sTmp1 := Memo.Lines.Strings[ind1];
                            iPos:=Pos('#39''', sTmp1);
                            if iPos>0 then sTmp1:=Copy(sTmp1,0,iPos-1) + Copy(sTmp1,iPos+4,255) ;
                            sTmp1 := Trim(Copy(sTmp1, Pos('''', sTmp1) + 1, 255));
                            sTmp1[Length(sTmp1)] := ' ';
                            TobTmp.PutValue('Libelle', sTmp1);
                            Break;
                            end;
                        Inc(ind1);
                        end;
                    end;
                end
            else if TobTmp.GetValue('Type') = 'THLabel' then
                TobTmp.Free
            else if TobTmp.GetValue('Type') = 'TBevel' then
                TobTmp.Free
            else if Copy(TobTmp.GetValue('Nom'), 0, 2) = 'XX' then
                TobTmp.Free
            else if TobTmp.Detail.Count = 0 then
                begin
                if not TobTmp.GetValue('Modif') then TobTmp.Free;
                end
            else if TobTmp.Detail[0].GetValue('Type') = 'DataField' then
                begin
                TobTmp.PutValue('Nom', TobTmp.Detail[0].GetValue('Nom'));
                if TobTmp.Detail[0].GetValue('Libelle') <> '' then
                    TobTmp.PutValue('Libelle', TobTmp.Detail[0].GetValue('Libelle'));
                TobTmp.PutValue('Modif', True);
                TobTmp.ClearDetail;
                end
            else if TobTmp.FillesCount(99) = 1 then
                begin
                TobTmp.PutValue('Nom', TobTmp.Detail[0].GetValue('Nom'));
                if TobTmp.Detail[0].GetValue('Libelle') <> '' then
                    TobTmp.PutValue('Libelle', TobTmp.Detail[0].GetValue('Libelle'));
                TobTmp.PutValue('Modif', True);
                TobTmp.ClearDetail;
                end;
            end;
        end
    else if (Pos('end', sTmp1) <> 0) and (Level > 1) then
        begin
        Dec(Level);
        Exit;
        end
    else
        begin
        sTmp1 := Memo.Lines.Strings[Ligne];
        sTmp2 := TobMere.GetValue('Type');
        if (Pos('GroupBox', sTmp2) <> 0) or
           (Pos('Panel', sTmp2) <> 0) or
           (Pos('Button', sTmp2) <> 0) or
           (Pos('TabSheet', sTmp2) <> 0) or
           (Pos('MenuItem', sTmp2) <> 0) then
            begin
            if (Pos(' Caption', sTmp1) <> 0) or (Pos(' Hint', sTmp1) <> 0) then
                begin
                sTmp1 := Trim(Copy(sTmp1, Pos('''', sTmp1) + 1, 255));
                sTmp1[Length(sTmp1)] := ' ';
                iPos:=Pos('#39''', sTmp1);
                if iPos>0 then sTmp1:=Copy(sTmp1,0,iPos-1) + Copy(sTmp1,iPos+4,255) ;
                TobMere.PutValue('Libelle', sTmp1);
                end;
            end
            else
            begin
            if (Pos('=', sTmp1) <> 0) then
                begin
                System.Delete(sTmp1, 1, Pos('''', sTmp1));
                sTmp1[Length(sTmp1)] := ' ';
                sTmp2 := Copy(sTmp1, 0, Pos('_', sTmp1) - 1);
                if sTmp2 = '' then
                    begin
                    sTmp1 := TobMere.GetValue('Nom');
                    sTmp2 := Copy(TobMere.GetValue('Nom'), 0, Pos('_', TobMere.GetValue('Nom')) - 1);
                    end;
                Table := mcd.getTable(Prefixetotable(sTmp2));
                FieldList := Table.Fields;
                FieldList.Reset();
                While FieldList.MoveNext do
                    begin
                    if (FieldList.Current as IFieldCOM).name = sTmp1 then
                        begin
                        TobTmp := TOB.Create('', TobMere, -1);
                        TobTmp.AddChampSup('Nom', True);
                        TobTmp.AddChampSup('Type', True);
                        TobTmp.AddChampSup('Libelle', True);
                        TobTmp.AddChampSup('Modif', True);
                        TobTmp.AddChampSup('Selected', True);
                        TobTmp.PutValue('Nom', Trim(sTmp1));
                        sTmp1 := Memo.Lines.Strings[Ligne];
                        System.Delete(sTmp1, Pos('=', sTmp1), 255);
                        TobTmp.PutValue('Type', Trim(sTmp1));
                        TobTmp.PutValue('Modif', True);
                        TobTmp.PutValue('Libelle', (FieldList.Current as IFieldCOM).Libelle);
                        if TobModif.FindFirst(['GOB_NOMCHAMP'], [TobTmp.GetValue('Nom')], False) <> nil then
                            TobTmp.PutValue('Selected', True)
                            else
                            TobTmp.PutValue('Selected', False);
                        if (FieldList.Current as IFieldCOM).Obligatoire then TobTmp.Free;
                        while (Pos('end', Memo.Lines.Strings[Ligne + 1]) = 0) do
                            Inc(Ligne);
                        Break;
                        end;
                    end;
                end;
            end;
        end;
    Inc(Ligne);
    end;
end ;
{$IFDEF GRC}
// construction du Treeview à partir de la tob mémoire des infos complémentaires
procedure TOF_PARAMCONFID.RTLoadTV(var Ligne, Level : integer; TobMere : TOB);
var i,NoOnglet,Panel : integer;
    TobMereOnglet,FilleTobParam,TobTmp,TobChampsProFille : TOB;
    TypeC,Code,Select,NomOnglet,Oblig: String;
begin
TobTmp := TOB.Create('', TobMere, -1);
RTAddTobTmp(TobTmp,'Pages','Pages','TPageControl');
Inc(Ligne);
Inc(Level);

//LoadTV(Ligne, Level, TobTmp);

TobMereOnglet:=TobTmp;
TobChampsProFille:=VH_RT.TobChampsProMul.detail[0];
for i := 0 to TobChampsProFille.Detail.Count-1 do
    begin
    FilleTobParam := TobChampsProFille.Detail[i];

    Code:=FilleTobParam.GetValue('RDE_CODE');
    NoOnglet:=FilleTobParam.GetValue('RDE_ONGLET');
    NomOnglet:=FilleTobParam.GetValue('RDE_NOMONGLET');
    Select := FilleTobParam.GetValue('RDE_CRITERESEL');
    Panel:=FilleTobParam.GetValue('RDE_PANEL');
    TypeC:=FilleTobParam.GetValue('RDE_TYPECHAMP');
    Oblig:=FilleTobParam.GetValue('RDE_OBLIGATOIRE');
    //if (  Panel <> 99 ) and ( code <> 'LIG' ) and ( code <> 'ESP' ) then
    if Panel = 99 then
        begin
        TobTmp := TOB.Create('', TobMereOnglet, -1);
        RTAddTobTmp(TobTmp,'ONGLETLIBRE'+IntToStr(NoOnglet),NomOnglet,'TTabSheet');
        TobMere:=TobTmp;
        end
    else
        if ( code <> 'LIG' ) and ( code <> 'ESP' ) and ( Oblig = '-') then
            begin
            TobTmp := TOB.Create('', TobMere, -1);
            if typeC='BOOLEAN' then typeC:='TCheckBox'
            else if typeC='COMBO' then typeC:='THValComboBox'
            else if typeC='DOUBLE' then typeC:='THNumEdit'
            else if typeC='MULTI' then typeC:='TListBox'
            else if (typeC='DATE') or (typeC='EDIT') then typeC:='THEdit'
            else if typeC='BLOB' then typeC:='TRichEdit'
            ;
            RTAddTobTmp(TobTmp,Code,RechDom('RTLIBCHAMPSLIBRES',Code,FALSE),typeC);
            end
    end;
TobMereOnglet.PutValue('Modif', False);
end;
{$ENDIF}

procedure TOF_PARAMCONFID.RTAddTobTmp(var TobTmp:Tob; Nom,Libelle,TypeC : String);
begin
TobTmp.AddChampSup('Nom', True);
TobTmp.AddChampSup('Type', True);
TobTmp.AddChampSup('Libelle', True);
TobTmp.AddChampSup('Modif', True);
TobTmp.AddChampSup('Selected', True);
TobTmp.PutValue('Nom', Nom);
TobTmp.PutValue('Libelle', Libelle);
TobTmp.PutValue('Type', TypeC);
TobTmp.PutValue('Selected', False);
TobTmp.PutValue('Modif', True);
if TobModif.FindFirst(['GOB_NOMCHAMP'], [TobTmp.GetValue('Nom')], False) <> nil then
    TobTmp.PutValue('Selected', True);
end;


procedure TOF_PARAMCONFID.ChargeGRPUSR(TobMere : TOB);
var
   TobTmpGrp, TobTmpUsr, TobTemp, TobTemp2 : TOB;
   ind1, ind2 : integer;
begin
for ind1 := 0 to TobGrp.Detail.Count - 1 do
    begin
    TobTemp := TobGrp.Detail[ind1];
    TobTmpGrp := TOB.Create('', TobMere, -1);
    TobTmpGrp.Dupliquer(TobTemp, False, True);
    if Pos(TobTmpGrp.GetValue('UG_GROUPE'), TobMere.GetValue('GOB_GRVISIBLE')) = 0 then
        TobTmpGrp.PutValue('GRPVISIBLE', ' ')
        else
        TobTmpGrp.PutValue('GRPVISIBLE', 'X');
    if Pos(TobTmpGrp.GetValue('UG_GROUPE'), TobMere.GetValue('GOB_GRENABLE')) = 0 then
        TobTmpGrp.PutValue('GRPENABLE', ' ')
        else
        TobTmpGrp.PutValue('GRPENABLE', 'X');
    for ind2 := 0 to TobUsr.Detail.Count - 1 do
        begin
        TobTemp2 := TobUsr.Detail[ind2];
        if TobTmpGrp.GetValue('UG_GROUPE') <> TobTemp2.GetValue('US_GROUPE') then Continue;
        TobTmpUsr := TOB.Create('', TobTmpGrp, -1);
        TobTmpUsr.Dupliquer(TobTemp2, False, True);
        if Pos(TobTmpUsr.GetValue('US_UTILISATEUR'), TobMere.GetValue('GOB_USVISIBLE')) = 0 then
            TobTmpUsr.PutValue('USRVISIBLE', ' ')
            else
            TobTmpUsr.PutValue('USRVISIBLE', 'X');
        if Pos(TobTmpUsr.GetValue('US_UTILISATEUR'), TobMere.GetValue('GOB_USENABLE')) = 0 then
            TobTmpUsr.PutValue('USRENABLE', ' ')
            else
            TobTmpUsr.PutValue('USRENABLE', 'X');
        end;
    end;
end;

procedure TOF_PARAMCONFID.RTDuplicClick(Sender: TObject);
var Question,CodeOrigine : String;
begin
if (CodeActif.Text = ClientGC) or (CodeActif.Text = InfosClient) then
    Question:='Confirmez-vous la duplication de la confidentialité prospects vers les clients ?'
else
    Question:='Confirmez-vous la duplication de la confidentialité clients vers les prospects ?';
if PGIAsk(Question, 'Duplication') <> mrYes then exit;
if (CodeActif.Text = ClientGC) then CodeOrigine:=ProspectGC
else if (CodeActif.Text = ProspectGC) then CodeOrigine:=ClientGC
else if (CodeActif.Text = InfosClient) then CodeOrigine:=InfosProspect
else if (CodeActif.Text = InfosProspect) then CodeOrigine:=InfosClient
;
ExecuteSQL ('delete from PARAMOBLIG where GOB_OBLIGATOIRE<>"X" and GOB_CODE="'+CodeActif.Text+'"');
ExecuteSQL ('insert into PARAMOBLIG (GOB_CODE,GOB_GRENABLE,GOB_GRVISIBLE,GOB_NOMCHAMP,GOB_OBLIGATOIRE,GOB_PREFIXE,GOB_PRIOCONFID,GOB_USENABLE1,GOB_USENABLE2,GOB_USVISIBLE1,GOB_USVISIBLE2) SELECT "'+CodeActif.Text+'",GOB_GRENABLE,GOB_GRVISIBLE,GOB_NOMCHAMP,GOB_OBLIGATOIRE,GOB_PREFIXE,GOB_PRIOCONFID,GOB_USENABLE1,GOB_USENABLE2,GOB_USVISIBLE1,GOB_USVISIBLE2 FROM PARAMOBLIG WHERE GOB_CODE="'+CodeOrigine+'" and GOB_OBLIGATOIRE<>"X"');
HValCB1Click(Sender);
end;

Initialization
  registerclasses ( [ TOF_PARAMCONFID ] ) ;
end.
