unit UTofParSuspect;

interface
uses  Controls,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOB,UTOF,
      vierge,SaisUtil,Messages,Windows
      ;

Const C_Objects    : integer = 0;
      C_suspect    : integer = 1;
      C_tiers      : integer = 2;
      C_compl      : integer = 3;
Type
TOF_Parsuspect = Class (TOF)
  private
  LesColonnes : string ;
  GS : THGRID ;
  TOBTableCorr,TOBL :tob;
  Function  ChercherChamps (prefixe, Code : String) : String;
  procedure RenseigneChamps;
  procedure SupprimeLigne(ARow:integer) ;
  procedure InsertLigne(var ARow:integer) ;
  procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
  procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  procedure GSComboPopUp(Sender : TObject);
  function ValeurDispo (CodeRech : String ; LaTable : String) : String;
  public
  procedure OnArgument(stArgument : String ) ; override ;
  procedure OnUpdate ; override ;
  procedure OnLoad ; override ;
  procedure OnClose ; override ;
  procedure OnNew ; override ;
  procedure OnDelete ; override ;
END ;

implementation
//C_Objects    : integer = 0;
//C_suspect     : integer = 1;
//C_tiers      : integer = 2;
//C_compl    : integer = 3;
//Table PARSUSPECTCOR    RSP
//RSP_CHSUSPECT	VARCHAR(35)
//RSP_LIBELLE	VARCHAR(35)
//RSP_CHTIERS	VARCHAR(35)
//RSP_CHCOMPL	VARCHAR(35)
//RSP_CBSUSPECT	COMBO
//RSP_CBTIERS	COMBO
//RSP_CBCOMPL	COMBO
/////////////////////////////////////////////////////////////////////////////


procedure TOF_Parsuspect.OnArgument (stArgument : String ) ;
BEGIN
  GS:= THGRID(GetControl('GRIDCOR'));
  GS.OnRowEnter:=GSRowEnter ;
  GS.OnRowExit:=GSRowExit ;
  GS.OnCellExit:=GSCellExit ;
  GS.OnCellEnter:=GSCellEnter;
  LesColonnes := 'RSP_CBSUSPECT;RSP_CBTIERS;RSP_CBCOMPL';
  GS.ColFormats[C_suspect]:='CB=RTLIBCHPLIBSUSPECTS||<<Aucun>>';
  GS.ColFormats[C_tiers]:='CB=GCZONELIBRETIE| AND (CC_CODE LIKE "CT%" OR CC_CODE LIKE "CC%" OR CC_CODE LIKE "CD%" OR CC_CODE LIKE "CM%" OR CC_CODE LIKE "CB%")|<<Aucun>>';
  GS.ColFormats[C_compl]:='CB=RTLIBCHAMPSLIBRES| AND (CC_CODE LIKE "CL%" OR CC_CODE LIKE "ML%"  OR CC_CODE LIKE "TL%" OR CC_CODE LIKE "DL%" OR CC_CODE LIKE "VL%" OR CC_CODE LIKE "BL%")|<<Aucun>>';
  AffecteGrid(GS,taModif) ;
  TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
END;

procedure TOF_Parsuspect.OnLoad;
begin
  inherited;
  TOBTableCorr:=tob.create('Correspondance',Nil,-1) ;
  TOBTableCorr.LoadDetailDB('PARSUSPECTCOR','RSP_CHSUSPECT','',nil,true);
  If TOBTableCorr.detail.count=0 then
  begin
    TOBL := Tob.create ('PARSUSPECTCOR',TOBTableCorr,-1) ;
    TOBL.initvaleurs;
  end;
  TOBTableCorr.PutGridDetail(GS,False,False,LesColonnes,false);
  if GS.valcombo<>nil then
  begin
    GS.valcombo.OnDropDown := GSComboPopup;
    GS.ShowCombo(GS.Col,GS.Row) ;
  end;
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS);
end;

procedure TOF_Parsuspect.OnUpdate ;
var TOBTemp :TOB;
begin
  PostMessage(GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB,  0) ;
  Application.processMessages;
  GS.Col := C_compl;GS.Row := GS.rowcount-1;
  GS.ShowCombo(GS.Col,GS.Row) ;
  if TOBTableCorr.Detail.Count > 0 then RenseigneChamps;
  TOBTemp := TOB.Create('maj', nil, -1);
  TOBTemp.LoadDetailDB ('PARSUSPECTCOR','RSP_CHSUSPECT','',nil,true);
  TOBTemp.DeleteDB (False);
  if (TOBTableCorr.Detail.Count > 1)
  or ((TOBTableCorr.Detail.Count=1) and (TOBTableCorr.Detail[0].getvalue('RSP_CHSUSPECT')<>'')) then
  begin
    TOBTemp.Dupliquer(TOBTableCorr,true,true,true);
    TOBTemp.InsertDB(nil,false);
  end;
  TOBTemp.free;
end ;

Procedure TOF_Parsuspect.RenseigneChamps ;
var i_ind :integer;
    CBSUSPECT,CHSUSPECT :string;
    CBTIERS,CHTIERS :string;
    CBCOMPL,CHCOMPL :string;
    TR :TOB;
begin
  for i_ind :=GS.RowCount-1 downto GS.fixedrows do
  begin
    if (trim (GS.cells[C_suspect,i_ind])='') or (GS.cells[C_suspect,i_ind]=TraduireMemoire('<<Aucun>>'))  then
    SupprimeLigne(i_ind);
  end;
  for i_ind := GS.fixedrows to GS.RowCount-1 do
  begin
    TOBL := TOB(GS.Objects[C_Objects,i_ind]); CBTIERS := '';
    if ( TOBL=nil) then continue;
    TOBL.InitValeurs;
    TOBL.GetLigneGrid(GS,i_ind,'FIXED;RSP_CBSUSPECT;RSP_CBTIERS;RSP_CBCOMPL' );
//    TOBL.putvalue ('RSP_LIBELLE',GS.cells[C_suspect,i_ind]);
    if (GS.Cells[C_tiers, i_ind] = '') or (GS.Cells[C_tiers, i_ind] = '<<Aucun>>') then
      TOBL.PutValue('RSP_LIBELLE', GS.Cells[C_compl, i_ind])
    else
      TOBL.PutValue('RSP_LIBELLE', GS.Cells[C_tiers, i_ind]);
    CBSUSPECT := TOBL.getvalue ('RSP_CBSUSPECT');
    CHSUSPECT := ChercherChamps('RSC',CBSUSPECT);
    TOBL.putvalue ('RSP_CHSUSPECT',CHSUSPECT);
    CBTIERS := TOBL.getvalue ('RSP_CBTIERS');
    CHTIERS := ChercherChamps('YTC',CBTIERS);
    TOBL.putvalue ('RSP_CHTIERS',CHTIERS);
    CBCOMPL := TOBL.getvalue ('RSP_CBCOMPL');
    CHCOMPL := ChercherChamps('RPR',CBCOMPL);
    TOBL.putvalue ('RSP_CHCOMPL',CHCOMPL);
  end;
  for i_ind :=GS.RowCount-1 downto GS.fixedrows do
  begin
    TOBL := TOB(GS.Objects[C_Objects,i_ind]);
    if ( TOBL=nil) then continue;
    CBSUSPECT := TOBL.getvalue ('RSP_CBSUSPECT');
    TR:=TOBTableCorr.FindFirst(['RSP_CBSUSPECT'],[CBSUSPECT],False) ;
    if (TR<>nil) and (TR<>TOBL) then SupprimeLigne(i_ind);
  end;
end;

Function TOF_Parsuspect.ChercherChamps (prefixe, Code : String) : String;
var champ :string;
    INum : integer;
begin
  champ := '';
  result:='';
  if Code = '' then exit;
  INum:=Ord(Code[3]);
  if INum < 65 then
     INum:=StrToInt(Code[3])
  else
     INum:=Ord(Code[3])-55;

  if copy(Code,1,2) = 'TL' then champ := prefixe + '_' + prefixe + 'LIBTEXTE'+IntToStr(INum)
  else if copy(Code,1,2) = 'VL' then champ := prefixe + '_' + prefixe + 'LIBVAL'+IntToStr(INum)
  else if copy(Code,1,2) = 'ML' then champ := prefixe + '_' + prefixe + 'LIBMUL'+IntToStr(INum)
  else if copy(Code,1,2) = 'BL' then champ := prefixe + '_' + prefixe + 'LIBBOOL'+IntToStr(INum)
  else if copy(Code,1,2) = 'DL' then champ := prefixe + '_' + prefixe + 'LIBDATE'+IntToStr(INum)
  else if copy(Code,1,2) = 'CC' then champ := 'YTC_TEXTELIBRE'+IntToStr(INum)
  else if copy(Code,1,2) = 'CM' then champ := 'YTC_VALLIBRE'+IntToStr(INum)
  else if copy(Code,1,2) = 'CT' then champ := 'YTC_TABLELIBRETIERS'+IntToHex(INum, 1)     //FQ 10687 
  else if copy(Code,1,2) = 'CB' then champ := 'YTC_BOOLLIBRE'+IntToStr(INum)
  else if copy(Code,1,2) = 'CD' then champ := 'YTC_DATELIBRE'+IntToStr(INum)
  else if copy(Code,1,2) = 'CL' then
  begin
    if Inum < 26 then         // table RSCLIB
      champ := prefixe + '_' + prefixe + 'LIBTABLE'+IntToStr(INum)
    else
      champ := prefixe + '_6' + prefixe + 'LIBTABLE'+IntToStr(INum);
  end;
  result := champ;
end;

procedure TOF_Parsuspect.InsertLigne(var ARow:integer) ;
begin
  PostMessage(GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB,  0) ;
  Application.processMessages;
  GS.CacheEdit; GS.SynEnabled := false;
  inc (Arow); GS.InsertRow (ARow);
  GS.Row := ARow;GS.Rows[ARow].Clear;
  GS.MontreEdit; GS.SynEnabled := true;
  TOBL := Tob.create ('PARSUSPECTCOR',TOBTableCorr,ARow-1) ;
  TOBL.InitValeurs;
  TOBL.PutLigneGrid(GS, ARow, false, false, LesColonnes);
  GS.ShowCombo(C_suspect,GS.Row) ;
end;

procedure TOF_Parsuspect.SupprimeLigne(ARow:integer) ;
begin
  if (GS.RowCount = 2) then
  begin
    TOB(GS.objects[0,ARow]).InitValeurs;
    TOB(GS.objects[0,ARow]).PutLigneGrid(GS, ARow, false, false, LesColonnes);
  end;
  if (GS.RowCount < 3) then Exit ;
  GS.CacheEdit; GS.SynEnabled := False;
  if (GS.objects[C_Objects,ARow]<>Nil) then TOB(GS.objects[0,ARow]).Free;
  GS.DeleteRow(ARow) ;
  GS.MontreEdit; GS.SynEnabled := True;
end ;

procedure TOF_Parsuspect.OnClose;
begin
  inherited;
  TOBTableCorr.cleardetail;
  TOBTableCorr.free;
end;

procedure TOF_Parsuspect.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
begin
FocusGrid := False;  ARow:=0;
if((Screen.ActiveControl = GS) or (Screen.ActiveControl.Owner = GS)) then
    BEGIN
    FocusGrid := True;
    ARow := GS.Row;
    END ;
Case Key of
    VK_INSERT : BEGIN
        if FocusGrid then
            BEGIN
            Key := 0;
            InsertLigne (ARow);
            END;
        END;
    VK_DOWN   : BEGIN
        if FocusGrid and (ARow=GS.RowCount-1) then
            BEGIN
            Key := 0;
            InsertLigne (ARow);
            END;
                END;
    VK_DELETE : BEGIN
        if (FocusGrid) then
            BEGIN
            Key := 0 ;
            SupprimeLigne (ARow) ;
            END ;
        END;
    END;
end;

procedure TOF_Parsuspect.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  GS.InvalidateRow(ou) ;
end;

procedure TOF_Parsuspect.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if (trim (GS.cells[C_suspect,ou])='') or (GS.cells[C_suspect,ou]=TraduireMemoire('<<Aucun>>')) then
     SupprimeLigne(ou);
  if not Cancel then GS.InvalidateRow(Ou);
end;

procedure TOF_Parsuspect.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
 if ( (GS.Col=C_tiers) and (GS.Cells [C_Compl,GS.Row]<>TraduireMemoire('<<Aucun>>')) and (GS.Cells [C_Compl,GS.Row]<>'') )
  or ( (GS.Col=C_Compl) and (GS.Cells [C_tiers,GS.Row]<>TraduireMemoire('<<Aucun>>')) and (GS.Cells [C_tiers,GS.Row]<>'') ) then
   GS.ValCombo.Enabled := False ;
   // on empêche la saisie dans les combo Tiers et compl si on a pas séléctionné un champ suspect
  if ((GS.Col = C_tiers) or (GS.Col = C_compl)) and (GS.Cells[C_suspect, GS.Row] = TraduireMemoire('<<Aucun>>')) then
    GS.ValCombo.Enabled := false;

end;

procedure TOF_Parsuspect.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if (ACol=C_suspect) and (GS.Cells [C_suspect,ARow]<>'') and (GS.objects[C_Objects,ARow]<>Nil) then
  TOB(GS.objects[C_Objects,ARow]).putvalue ('RSP_CBSUSPECT',GS.CellValues[ACol, ARow]);
  // on alimente la Tob
  if (ACol=C_tiers) and (GS.Cells [C_tiers,ARow]<>'') and (GS.objects[C_Objects,ARow]<>Nil) then      //FQ 10494
    TOB(GS.objects[C_Objects,ARow]).putvalue ('RSP_CBTIERS',GS.CellValues[ACol, ARow]);
  if (ACol=C_compl) and (GS.Cells [C_compl,ARow]<>'') and (GS.objects[C_Objects,ARow]<>Nil) then      //FQ 10494
    TOB(GS.objects[C_Objects,ARow]).putvalue ('RSP_CBCOMPL',GS.CellValues[ACol, ARow]);

end;

procedure TOF_Parsuspect.GSComboPopUp(Sender: TObject);
var St_CBSuspect,St_Value,St_Text,St_Plus : string;
    FCombo : THValComboBox;
    StrCBSuspectComplet : String;

begin
  FCombo := (Sender as THValComboBox);
  St_Plus := '';
  St_Value := FCombo.Value;
  St_Text := FCombo.Text;

  StrCBSuspectComplet   := TOB(GS.objects[C_Objects,GS.Row]).getvalue ('RSP_CBSUSPECT');
  St_CBSuspect          := copy (StrCBSuspectComplet, 1, 2);

  if (GS.Col = C_suspect) then        //FQ10494
    St_Plus             := ValeurDispo(St_Value, 'SUSPECT');

  if (GS.Col =C_tiers) then
  begin
    if (St_CBSuspect = '') then
    St_Plus := ' AND (CC_CODE LIKE "CT%" OR CC_CODE LIKE "CC%" OR CC_CODE LIKE "CD%" OR CC_CODE LIKE "CM%" OR CC_CODE LIKE "CB%") '
        // FQ 10494 on ajoute en plus de la séléction, la valeur de retour de ValeurDispo afin d'écarter les éléments utilisés
    else if (St_CBSuspect = 'CL') then St_Plus := ' AND CC_CODE LIKE "CT%" ' + ValeurDispo('CT', 'TIERS')
    else if (St_CBSuspect = 'TL') then St_Plus := ' AND CC_CODE LIKE "CC%" ' + ValeurDispo('CC', 'TIERS')
    else if (St_CBSuspect = 'DL') then St_Plus := ' AND CC_CODE LIKE "CD%" ' + ValeurDispo('CD', 'TIERS')
    else if (St_CBSuspect = 'VL') then St_Plus := ' AND CC_CODE LIKE "CM%" ' + ValeurDispo('CM', 'TIERS')
    else if (St_CBSuspect = 'BL') then St_Plus := ' AND CC_CODE LIKE "CB%" ' + ValeurDispo('CB', 'TIERS');
  end;
  if (GS.Col =C_compl) then
  begin
    if (St_CBSuspect = '') then
    St_Plus := ' AND (CC_CODE LIKE "CL%" OR CC_CODE LIKE "ML%"  OR CC_CODE LIKE "TL%" OR CC_CODE LIKE "DL%" OR CC_CODE LIKE "VL%" OR CC_CODE LIKE "BL%") '
    // Les tables libres prospect à 6 caractères commencent à CLQ -> CLZ
    else if (St_CBSuspect = 'CL') and (StrCBSuspectComplet >= 'CLQ') then      //FQ 10523
//      St_Plus           := ' AND (CC_CODE LIKE "CL%" OR CC_CODE LIKE "ML%") ' + ValeurDispo('CL', 'COMPL') + ValeurDispo('ML', 'COMPL')
    //FQ 10770   TJA 27/02/2006
      St_Plus           := ' AND (CC_CODE LIKE "ML%") ' + ValeurDispo('ML', 'COMPL')
    else if (St_CBSuspect = 'CL') and (StrCBSuspectComplet < 'CLQ') then
      St_Plus           := ' AND (CC_CODE LIKE "CL%" AND CC_CODE < "CLQ") ' + ValeurDispo('CL', 'COMPL')
    else St_Plus        := ' AND CC_CODE LIKE "'+St_CBSuspect+'%" ' + ValeurDispo(St_CBSuspect, 'COMPL');
  end;
  FCombo.Plus := St_Plus;
  FCombo.Value := St_Value;
  FCombo.Text := St_Text;
end;

procedure TOF_Parsuspect.OnNew;
var ARow : integer;
begin
  inherited;
  ARow := GS.Row;
  InsertLigne (ARow);
end;

procedure TOF_Parsuspect.OnDelete;
var ARow : integer;
begin
  inherited;
  ARow := GS.Row;
  SupprimeLigne (ARow) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/08/2007
Modifié le ... : 30/08/2007
Description .. : Alimente une chaine pour la propriété PLUS de la combo, 
Suite ........ : afin d'extraire les éléments déjà utilisés dans une 
Suite ........ : précédente correspondance
Mots clefs ... : 
*****************************************************************}
function TOF_Parsuspect.ValeurDispo(CodeRech : String ; LaTable : String): String;
var
  i                     : integer;
  LePlus                : String;
  StrParSuspect         : String;

begin
  result                := '';
  LePlus                := '';

  if UpperCase_(LaTable) = 'SUSPECT' then
  begin
    For i := 0 to TOBTableCorr.Detail.Count -1 do
      if (TOBTableCorr.Detail[i].GetValue('RSP_CBSUSPECT') <> CodeRech) or (CodeRech = '') then
        LePlus          := LePlus + ' AND CC_CODE <> "' + TOBTableCorr.Detail[i].GetValue('RSP_CBSUSPECT') + '"';
  end
  else
  begin                                                  //FQ 10494  on enlève des listes les éléments déjà utilisés
    if LaTable = 'TIERS' then
      StrParSuspect     := 'RSP_CBTIERS'
    else
      StrParSuspect     := 'RSP_CBCOMPL';

    for i := 0 to TOBTableCorr.Detail.Count -1 do
      if (Copy(TOBTableCorr.Detail[i].GetValue(StrParSuspect), 1, length(CodeRech)) = CodeRech) and (i <> GS.Row-1)  then
        LePlus          := LePlus + ' AND CC_CODE <> "' + TOBTableCorr.Detail[i].GetValue(StrParSuspect) + '"';
  end;

  Result                := LePlus;

end;

Initialization
registerclasses([TOF_Parsuspect]) ;

end.
