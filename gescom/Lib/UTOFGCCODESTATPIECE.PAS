{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/12/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCCODESTATPIECE ()
Mots clefs ... : TOF;GCCODESTATPIECE
*****************************************************************}
Unit UTOFGCCODESTATPIECE ;

Interface

Uses UTOF, HCtrls, UTOB, Classes, Graphics, Grids,
{$IFNDEF EAGLCLIENT}
     DBTables,
{$ENDIF}
     HEnt1, HMsgBox, SysUtils, SAISUTIL, Vierge, Forms,
     Windows, Controls, M3FP ;

Type
  TOF_GCCODESTATPIECE = Class (TOF)
    Private
    LesColonnes,TempCodeTable,CodeTable : string ;
    Col_Mov,Col_Code,Col_CodeEntier,Col_Lib,Col_Abr,Col_Type : Integer;
    GS : THGRID ;
    TobToDelete : Tob ;
    procedure InsertLigne(Ligne:integer) ;
    procedure SupprimeLigne(Ligne:integer) ;
    procedure SuppLigneVide;
    function  CodeUtiliseDansParPiece : boolean ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GSEnter(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure ExtractCode;
    procedure ChargeGrille;
    function  ValideModif(CodeTabLib:String;Question:boolean=True) : boolean;
    procedure EnregistrerGrille;
    function  ExisteCode(Code : String; Ligne : Integer) : boolean;
    procedure ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure CodeGras ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    Public
    Action   : TActionFiche ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    Procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_GCCODESTATPIECE.InsertLigne(Ligne:integer) ;
begin
GS.InsertRow(Ligne) ;
GS.Row:=Ligne; GS.Col:=Col_Code;
GS.Cells[Col_Type,GS.Row]:='GTL';
GS.Cells[Col_CodeEntier,GS.Row]:=CodeTable;
end;

procedure TOF_GCCODESTATPIECE.OnNew ;
var ACol,ARow : integer;
    Cancel : boolean;
begin
  Inherited ;
  if GS.RowCount=2 then GS.Cells[Col_Type,1]:='GTL';
  Cancel:=False; ACol:=GS.Col; ARow:=GS.Row;
  GSCellExit(GS,ACol,ARow,Cancel);
  if Not Cancel then GSRowExit(GS,ARow,Cancel,False);
  if Cancel then exit;
  InsertLigne(GS.Row);
end ;

function TOF_GCCODESTATPIECE.CodeUtiliseDansParPiece : boolean ;
var  SQL,CodePiece : string ;
begin
  Result:=False;
  CodePiece:=GS.Cells[Col_CodeEntier,GS.Row];
  SQL:='Select GPP_NATUREPIECEG from PARPIECE Where '+
       '(GPP_PIECETABLE1="'+CodeTable+'" OR GPP_PIECETABLE2="'+CodeTable+'" OR GPP_PIECETABLE3="'+CodeTable+'")'+
  ' AND (GPP_CODPIECEDEF1="'+CodePiece+'" OR GPP_CODPIECEDEF2="'+CodePiece+'" OR GPP_CODPIECEDEF3="'+CodePiece+'")';
  if ExisteSQL(SQL) then Result:=True;
end ;

procedure TOF_GCCODESTATPIECE.SupprimeLigne(Ligne:integer) ;
begin
if Ligne<=0 then Exit ;
if GS.RowCount<2 then Exit ;
if GS.RowCount=2 then GS.RowCount:=3;
if GS.objects[0,Ligne]<>Nil then
  begin
  if GS.Cells[Col_Code,Ligne]='' then TOB(GS.objects[0,Ligne]).Free
  else
    begin
    if Not CodeUtiliseDansParPiece then TOB(GS.objects[0,Ligne]).ChangeParent( TobToDelete,-1)
    else
      begin
      PGIInfo('Cette ligne ne peut pas être supprimée car cette valeur est utilisée dans le paramétrage des pièces !',Ecran.Caption);
      exit;
      end;
    end;
  end;
GS.DeleteRow(Ligne) ;
if Ligne=1 then
  begin
  GS.Row:=Ligne;
  if GS.Cells[Col_Code,Ligne]='' then GS.Col:=Col_Code else GS.Col:=Col_Lib;
  end
else
  begin
  GS.Row:=Ligne-1;
  if GS.Cells[Col_Code,Ligne-1]='' then GS.Col:=Col_Code else GS.Col:=Col_Lib;
  end;
end;

procedure TOF_GCCODESTATPIECE.OnDelete ;
begin
  Inherited ;
SupprimeLigne(GS.Row);
end ;

procedure TOF_GCCODESTATPIECE.SuppLigneVide;
var Ligne : integer;
begin
For Ligne:=GS.RowCount-1 downto 1 Do
  begin
  if GS.Cells[Col_Code,Ligne]='' then SupprimeLigne(Ligne);
  end;
end;

procedure TOF_GCCODESTATPIECE.OnUpdate ;
begin
  Inherited ;
ValideModif(CodeTable,False);
end ;

procedure TOF_GCCODESTATPIECE.OnLoad ;
begin
  Inherited;
if CodeTable='' then
  begin
  if THValComboBox(GetControl('_TABLELIBRE')).Items.Count<>0 then THValComboBox(GetControl('_TABLELIBRE')).ItemIndex:=0;
  CodeTable:=THValComboBox(GetControl('_TABLELIBRE')).Value;
  end
else THValComboBox(GetControl('_TABLELIBRE')).Value:=CodeTable;
ChargeGrille;
// Permet d'afficher en gras le code de la première ligne du Grid 
GS.SetFocus;
end ;

procedure TOF_GCCODESTATPIECE.OnArgument (S : String ) ;
var i : integer ;
    St,Nam,Critere,ChampMul,ValMul : string ;
begin
  Inherited ;
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  if Critere<>'' then
    begin
    i:=pos('=',Critere);
    if i<>0 then
      begin
      ChampMul:=copy(Critere,1,i-1);
      ValMul:=copy(Critere,i+1,length(Critere));
      if ChampMul='ACTION' then
        begin
        if ValMul='CREATION' then
          begin
          Action:=taCreat;
          THValComboBox(GetControl('_TABLELIBRE')).Enabled := False;
          end
          else if ValMul='MODIFICATION' then Action:=taModif
            else if ValMul='CONSULTATION' then Action:=taConsult;
        end;
      if ChampMul='YX_CODE' then CodeTable:=ValMul;
      end;
    end;
until  Critere='' ;
LesColonnes :='MOV;CODE;YX_LIBELLE;YX_ABREGE;YX_CODE;YX_TYPE';
GS:=THGRID(GetControl('G'));
GS.OnEnter:=GSEnter;
GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnRowEnter:=GSRowEnter ;
GS.OnRowExit:=GSRowExit ;
GS.GetCellCanvas:= CodeGras;
GS.PostDrawCell:= ColTriangle;
GS.ColCount:=1; i:=0;
St:=LesColonnes ;
Col_Mov:=-1; Col_Code:=-1; Col_Lib:=-1; Col_Abr:=-1; Col_CodeEntier:=-1;
Repeat
   Nam:=ReadTokenSt(St) ;
   if Nam<>'' then
     begin
     if Nam='MOV' then
       begin
       if i<>0 then GS.ColCount:=GS.ColCount+1;
       Col_Mov:=i; GS.ColWidths[Col_Mov]:=12;
       end
       else if Nam='CODE' then
         begin
         if i<>0 then GS.ColCount:=GS.ColCount+1;
         Col_Code:=i; GS.ColWidths[Col_Code]:=50; GS.ColLengths[Col_Code]:=3;
         GS.ColFormats[Col_Code]:='UPPER';
         end
         else if Nam='YX_LIBELLE' then
           begin
           if i<>0 then GS.ColCount:=GS.ColCount+1;
           Col_Lib:=i; GS.ColWidths[Col_Lib]:=150; GS.ColLengths[Col_Lib]:=70;
           end
           else if Nam='YX_ABREGE' then
             begin
             if i<>0 then GS.ColCount:=GS.ColCount+1;
             Col_Abr:=i; GS.ColWidths[Col_Abr]:=80; GS.ColLengths[Col_Abr]:=17;
             end
             else if Nam='YX_CODE' then
               begin
               if i<>0 then GS.ColCount:=GS.ColCount+1;
               Col_CodeEntier:=i; GS.ColWidths[Col_CodeEntier]:=0; GS.ColLengths[Col_CodeEntier]:=-1;
               end
               else if Nam='YX_TYPE' then
                 begin
                 if i<>0 then GS.ColCount:=GS.ColCount+1;
                 Col_Type:=i; GS.ColWidths[Col_Type]:=-1; GS.ColLengths[Col_Type]:=-1;
                 end;
       Inc(i);
       end;
Until ((St='') or (Nam='')) ;
if Col_Mov<>-1 then GS.FixedCols:=1;
AffecteGrid(GS,Action) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
end ;

procedure TOF_GCCODESTATPIECE.OnClose ;
begin
  Inherited ;
ValideModif(CodeTable);
end ;

Procedure TOF_GCCODESTATPIECE.OnCancel ;
begin
inherited ;
TobToDelete.free ;  TobToDelete:=Nil ;
LaTob.free ; LaTob:=Nil ;
ChargeGrille;
end ;

procedure TOF_GCCODESTATPIECE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid,Cancel : Boolean;
    ARow,ACol : Longint;
BEGIN
if Action=taConsult then Exit ;
FocusGrid := False;  ARow:=0;
if (Screen.ActiveControl = GS) then
    BEGIN
    FocusGrid := True;
    ARow := GS.Row;
    END ;
Case Key of
    VK_F10    : if FocusGrid then ValideModif(CodeTable,False);
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow) ;
                    END ;
                END;
    VK_DOWN   : BEGIN
                if FocusGrid then
                    BEGIN
                    if ARow=GS.RowCount-1 then
                      begin
                      if GS.RowCount=2 then GS.Cells[Col_Type,1]:='GTL';
                      Cancel:=False; ACol:=GS.Col; ARow:=GS.Row;
                      GSCellExit(GS,ACol,ARow,Cancel);
                      if Not Cancel then GSRowExit(GS,ARow,Cancel,False);
                      if Cancel then exit;
                      InsertLigne (GS.RowCount);
                      end;
                    END ;
                END;
    END;
END;

procedure TOF_GCCODESTATPIECE.GSEnter(Sender: TObject);
Var ACol,ARow : integer;
    Temp : Boolean;
begin
if Action=taConsult then Exit ;
if GS.Cells[Col_Type,1]='' then
  begin
  GS.Cells[Col_Type,1]:='GTL';
  GS.Cells[Col_CodeEntier,1]:=CodeTable;
  end;
ACol:=GS.Col; ARow:=GS.Row; Temp:=False;
GSCellEnter(GS,ACol,ARow,Temp);
end;

procedure TOF_GCCODESTATPIECE.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if GS.Col=Col_Code then
  begin
  if GS.Cells[Col_Code,GS.Row]<>'' then GS.Col:=Col_Lib;
  end;
end;

procedure TOF_GCCODESTATPIECE.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if (ACol=Col_Code) then
  begin
  if Not ExisteCode(GS.Cells[Col_Code,ARow],ARow) then
    GS.Cells[Col_CodeEntier,ARow]:=TempCodeTable+GS.Cells[ACol,ARow]
  else
    begin
    PGIInfo('Le code que vous avez saisi existe déjà. Vous devez le modifier',Ecran.Caption);
    GS.Cells[Col_Code,ARow]:=''; Cancel:=True;
    end;
  end
else if (ACol=Col_Lib) then
       begin
       if GS.Cells[Col_Abr,ARow]='' then GS.Cells[Col_Abr,ARow]:=copy(GS.Cells[Col_Lib,ARow],1,17);
       end;
end;

procedure TOF_GCCODESTATPIECE.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if Action=taConsult then Exit ;
GS.InvalidateRow(Ou);
end;

procedure TOF_GCCODESTATPIECE.GSRowExit(Sender: TObject; Ou: Integer;
                                            var Cancel: Boolean; Chg: Boolean);
begin
if Action=taConsult then Exit ;
if Col_Lib=-1 then exit;
if GS.Cells[Col_Lib,Ou]='' then
  begin
  if GS.Cells[Col_Code,Ou]<>'' then
    begin
    PGIInfo('Vous devez renseigner un libellé.',Ecran.Caption);
    GS.Col:=Col_Lib; Cancel:=True;
    end;
  end
else
  begin
  if GS.Cells[Col_Code,Ou]='' then
    begin
    PGIInfo('Vous devez renseigner un code.',Ecran.Caption);
    GS.Col:=Col_Code; Cancel:=True;
    end;
  end;
if not Cancel then GS.InvalidateRow(Ou);
end;

Procedure TOF_GCCODESTATPIECE.ExtractCode;
Var i : integer;
begin
For i:=0 to LaTOB.Detail.Count-1 Do
  begin
  LaTOB.Detail[i].AddChampSup('CODE',False);
  LaTOB.Detail[i].PutValue('CODE',Copy(LaTOB.Detail[i].GetValue('YX_CODE'),4,3));
  end;
end;

Procedure TOF_GCCODESTATPIECE.ChargeGrille;
Var Q : TQuery;
begin
If ValideModif(CodeTable) then
  begin
  TempCodeTable:=CodeTable;
  LaTob:=tob.create('_CHOIXEXT',Nil,-1) ;
  Q:=OpenSQL('SELECT YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE FROM CHOIXEXT WHERE YX_TYPE="GTL" AND YX_CODE LIKE "'+CodeTable+'%" ORDER BY YX_CODE',True);
  if Not Q.EOF then Latob.LoadDetailDB('CHOIXEXT','','',Q,false,true);
  Ferme(Q);
  If LaTob.detail.count=0 then Tob.create ('CHOIXEXT',LaTob,-1);
  ExtractCode;
  Latob.PutGridDetail(GS,True,True,LesColonnes,True);
  LaTob.GetGridDetail(GS,GS.rowcount-1,'CHOIXEXT',LesColonnes);
  LaTob.SetAllModifie (false);
  if Col_Code<>-1 then GS.Cells[Col_Mov,0]:='';
  if Col_Code<>-1 then GS.Cells[Col_Code,0]:='Code';
  TobToDelete:=tob.create('_CHOIXEXT à supprimer',Nil,-1) ;
  end;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
end;

function TOF_GCCODESTATPIECE.ValideModif(CodeTabLib:String;Question:boolean=True) : boolean;
Var Recharger,Cancel : boolean;
    Reponse,ACol,ARow : integer;
begin
Result:=False; Recharger:=False;
if CodeTabLib='' then exit;
if LaTob=nil then begin Result:=True; exit; end;
Cancel:=False; ACol:=GS.Col; ARow:=GS.Row;
GSCellExit(GS,ACol,ARow,Cancel);
if Not Cancel then GSRowExit(GS,ARow,Cancel,False);
if Cancel then exit;
if GS.rowcount > 1 then
  begin
  LaTob.GetGridDetail(GS,GS.rowcount-1,'CHOIXEXT',LesColonnes) ;
  SuppLigneVide;
  end;
if Question then
  begin
  if (LaTob.IsOneModifie) or (TobToDelete.detail.count > 0) then
      begin
      Reponse:= PGIAsk('Voulez-vous enregistrer les modifications ?',Ecran.Caption);
      if Reponse=mrYes then
        begin
        Recharger:=True; EnregistrerGrille;
        end
        else if Reponse=mrNo then
          begin
          Recharger:=True;
          end;
      end;
  end
else EnregistrerGrille;
if Not Recharger and (CodeTabLib=TempCodeTable) then exit;
LaTob.free ; Latob:=nil;
TobToDelete.free; TobToDelete:=nil;
Result:=True;
end;

procedure TOF_GCCODESTATPIECE.EnregistrerGrille;
begin
// suppressions
TobToDelete.DeleteDB (False) ;
TobToDelete.ClearDetail ;
// MAJ
LaTob.InsertOrUpdateDB(False) ;
LaTob.SetAllModifie(false);
end;

function TOF_GCCODESTATPIECE.ExisteCode(Code : String; Ligne : Integer) : boolean;
var i : integer;
begin
Result:=False;
if Code='' then exit;
For i:=1 to GS.RowCount-1 Do
  begin
  if (GS.Cells[Col_Code,i]=Code) and (i<>Ligne) then
    begin Result:=True; Break; end;
  end;
end;

procedure TOF_GCCODESTATPIECE.ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
  Arect:=GS.CellRect(Acol,Arow) ;
  Canvas.Brush.Color := GS.FixedColor;
  Canvas.FillRect(ARect);
    if (ARow = GS.row) then
       begin
       Canvas.Brush.Color := clBlack ;
       Canvas.Pen.Color := clBlack ;
       Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
       Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
       Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
       if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
       end ;
  end;
end;

procedure TOF_GCCODESTATPIECE.CodeGras ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);
begin
if (ACol = Col_Code) and (ARow>0) then
  begin
  Canvas.Font.Style := [fsBold];
  end;
end;

procedure AGLMAJGrille( parms: array of variant; nb: integer ) ;
var  F : TForm;
     LaTOF : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFVierge) then LaTOF:=TFVierge(F).LaTOF else exit;
if (LaTOF is TOF_GCCODESTATPIECE) then
  begin
  TOF_GCCODESTATPIECE(LaTOF).CodeTable:=Parms[1];
  TOF_GCCODESTATPIECE(LaTOF).ChargeGrille;
  end
else exit;
end;

Initialization
  registerclasses ( [ TOF_GCCODESTATPIECE ] ) ;
  RegisterAglProc( 'MAJGrille', TRUE , 1, AGLMAJGrille);
end.

