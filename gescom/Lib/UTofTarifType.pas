unit UTofTarifType;

interface                                                                
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,TarifUtil,DBTables,M3FP,Fe_Main,HDB,voirtob;

Type
    TOF_TarifType = Class(TOF)
     private
        GridTyp : THGRID ;
        TobTyp,TobTypLigne,TOBToDelete: Tob ;
        NewLigne: Boolean ;
        Col_Mov, colCode, colLibelle,  colDevise, colCoef,colEtablis: Integer ;
        LesColonnes: String ;

        procedure GSElipsisClick(Sender: TObject);
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSEnter(Sender: TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure ChargePeriode ;
        procedure FormateColSaisie (ACol,ARow : Longint ) ;

        //Function VerifCode(Code: String): boolean ;
        Function VerifCombo(Champ,Code: String): boolean ;
        Function VerifEnreg: Boolean ;

        function  TypeUtiliseDansTarif : boolean ;

        Function  EstRempli( Lig : integer) : Boolean ;
        procedure ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
        procedure CodeGras ( ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);


     public
         Action   : TActionFiche ;
         procedure OnArgument (Arguments : String ) ; override ;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         Procedure OnNew  ; override ;
         Procedure OnDelete ; override ;
         Procedure OnCancel ; override ;
         Procedure OnClose ; override ;

     END ; 

implementation

Procedure TOF_TarifType.OnArgument(Arguments:string) ;
var St,S,NomCol :String ;
i: Integer ;
Begin
inherited ;
St:=Arguments ;
Action:=taModif ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;
LesColonnes:='MOV;GFT_CODETYPE;GFT_LIBELLE;GFT_DEVISE;GFT_COEFFICIENT;GFT_ETABLISREF' ;
GRIDTYP:=THGRID(GetControl('GRIDTYP'));
GRIDTYP.OnElipsisClick:=GSElipsisClick  ;
GRIDTYP.OnDblClick:=GSElipsisClick ;
GRIDTYP.OnCellEnter:=GSCellEnter ;
GRIDTYP.OnCellExit:=GSCellExit ;
GRIDTYP.OnRowExit:=GSRowExit ;
GRIDTYP.OnRowEnter:=GSRowEnter ;
GRIDTYP.OnEnter:=GSEnter ;
GRIDTYP.GetCellCanvas:= CodeGras;
GRIDTYP.PostDrawCell:= ColTriangle;
GRIDTYP.ColCount:=1; i:=0;
S:=LesColonnes ;
Col_Mov:=-1; colCode :=-1; colLibelle:=-1;  colDevise:=-1 ; colCoef:=-1 ; colEtablis:=-1 ;
Repeat
   NomCol:=ReadTokenSt(S) ;
   if NomCol<>'' then
     begin
     if NomCol='MOV' then
       begin
       if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1;
       Col_Mov:=i; GRIDTYP.ColWidths[Col_Mov]:=10;
       end
       else if NomCol='GFT_CODETYPE' then
         begin
         if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1;
         colCode:=i; GRIDTYP.ColWidths[colCode]:=50; GRIDTYP.ColLengths[colCode]:=6;
         GRIDTYP.ColFormats[colCode]:='UPPER';
         end
         else if NomCol='GFT_LIBELLE' then
           begin
           if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1;
           colLibelle:=i; GRIDTYP.ColWidths[colLibelle]:=150; GRIDTYP.ColLengths[colLibelle]:=70;
           end
           else if NomCol='GFT_DEVISE' then
             begin
             if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1;
             colDevise:=i ;
             GRIDTYP.ColFormats[colDevise]:='UPPER';
             end
             else if NomCol='GFT_COEFFICIENT' then
               begin
               if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1;
               colCoef:=i ;
               end
               else if NomCol='GFT_ETABLISREF' then
               begin
               if i<>0 then GRIDTYP.ColCount:=GRIDTYP.ColCount+1 ;
               colEtablis:=i ; GRIDTYP.ColFormats[colEtablis]:='UPPER';
               GRIDTYP.ColWidths[colEtablis]:=150;
               end ;
       Inc(i);
       end;
Until ((St='') or (NomCol='')) ;
if Col_Mov<>-1 then GRIDTYP.FixedCols:=1;

AffecteGrid(GRIDTYP,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GRIDTYP) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
End ;

Procedure TOF_TarifType.OnLoad ;
Begin
inherited ;
ChargePeriode ;
// Permet d'afficher en gras le code de la première ligne du Grid
GRIDTYP.SetFocus ;
End ;

Procedure TOF_TarifType.OnUpdate ;
Begin
inherited ;
SetControlEnabled('BInsert',True) ;
SetControlEnabled('BDelete',True) ;
TobTyp.GetGridDetail(GRIDTYP,GridTyp.rowcount-1,'TARIFTYPMODE',LesColonnes);
If not VerifEnreg then exit ;
TobTyp.InsertOrUpdateDB(False) ;
TobTyp.SetAllModifie(false);
NewLigne:=False ;
End ;

Procedure TOF_TarifType.OnNew ;
Begin
inherited ;
NewLigne:=True ;
GRIDTYP.InsertRow(GRIDTYP.Row) ;
GRIDTYP.Row:=GRIDTYP.Row-1 ;
GRIDTYP.Col:=ColCode ;
//GRIDTYP.ColLengths[colCode]:=6 ;
GRIDTYP.SetFocus ;
SetControlEnabled('BInsert',False) ;
SetControlEnabled('BDelete',False) ;
end ;

Procedure TOF_TarifType.OnDelete  ;
begin
inherited ;
if GRIDTYP.Row<=0 then Exit ;
if GRIDTYP.RowCount<2 then Exit ;
if GRIDTYP.RowCount=2 then GRIDTYP.RowCount:=3;
if TypeUtiliseDansTarif then
      begin
       PGIInfo('Un tarif de ce type est utilisé,Vous ne pouvez pas le supprimer !',Ecran.Caption);
       exit;
      end else
      begin
       if HShowMessage('1;Période d''application;Confirmer vous la suppression de l''enregistrement ?;Q;YNC;N;C;','','')<>mrYes then exit;
       ExecuteSQL('DELETE FROM TARIFTYPMODE WHERE GFT_CODETYPE="'+GRIDTYP.Cells[ColCode,GRIDTYP.Row]+'"') ;
       GRIDTYP.DeleteRow(GRIDTYP.row) ;
      end;
end;

Procedure TOF_TarifType.OnCancel ;
begin
inherited ;
SetControlEnabled('BInsert',True) ;
SetControlEnabled('BDelete',True) ;
If GRIDTYP.RowCount-1>TobTyp.Detail.count then
   if NewLigne then
     begin
     GRIDTYP.DeleteRow(1) ;
     NewLigne:=False ;
     end ;
OnLoad ;
end ;

Procedure TOF_TarifType.OnClose ;
Var i:Integer ;
TOBTypeSav: TOB ;
Compare: integer ;
Begin
inherited ;
Compare:=0 ;
TOBTypeSav:= TOB.Create('_TARIFTYPMODE',NIL,-1);
TOB.Create ('TARIFTYPMODE',nil,-1) ;
TOBTypeSav.Dupliquer (TobTyp, True, True);
LastError:=0;
if GridTyp.rowcount > 1 then TobTyp.GetGridDetail(GridTyp,GridTyp.RowCount,'TARIFTYPMODE',LesColonnes);
if TOBTypeSav.Detail.Count=TobTyp.Detail.count then
  begin
  For i:=0 to TobTyp.Detail.count-1 do
    begin
    if TobTyp.Detail.count>1 then
      Compare:=CompareTOB (TobTyp.Detail[i],TOBTypeSav.Detail[i],LesColonnes);
    if Compare<>0 then  break ;
    end ;
  end else Compare:=10 ;
if Compare<>0  then
  begin
  If PGIAsk('Confirmez-vous l''abandon des modifications?',Ecran.Caption)=mrNo then
      begin
      LastError:=-1;
      Exit;
      end else
      begin
      LastError:=0;
      exit;
      end;
  end;
TOBTypeSav.free ; TOBTypeSav:=nil;
TobTyp.free ; TobTyp:=nil;
end;

Function TOF_TarifType.VerifEnreg: Boolean ;
var ARow: Integer ;
begin
ARow:=GridTyp.Row ;
Result:=True ;
If VerifCombo('CODETYP',GridTyp.Cells[ColCode,GridTyp.Row]) and NewLigne then
   begin
    PGIBox('Le type de tarif que vous avez saisi existe déjà. Vous devez le modifier',Ecran.Caption) ;
    Result:=False ;
    GridTyp.Col := colCode;
    exit ;
   end ;
If VerifCombo('DEVISE',GridTyp.Cells[ColDevise,GridTyp.Row]) then
   begin
    PGIBox('Le code devise n''existe pas. Vous devez le modifier',Ecran.Caption) ;
    Result:=False ;
    GridTyp.Col := colDevise;
    exit ;
   end ;
If VerifCombo('ETABLISSEMENT',GridTyp.Cells[ColEtablis,GridTyp.Row]) then
   begin
    PGIBox('Le code établissement n''existe pas. Vous devez le modifier',Ecran.Caption) ;
    Result:=False ;
    GridTyp.Col := ColEtablis;
    exit ;
   end ;
if(GRIDTYP.Cells[ColCode,Arow]='') then
    begin
    PGIBox('Le code est obligatoire',Ecran.Caption) ;
    GRIDTYP.Col:=ColCode ;
    end ;
if (GRIDTYP.Cells[ColLibelle,Arow]='')  then
   begin
   PGIBox('Le libelle est obligatoire',Ecran.Caption) ;
   GRIDTYP.Col:=ColLibelle ;
   end ;
if (GRIDTYP.Cells[ColEtablis,Arow]='')  then
   begin
   PGIBox('L''établissement de référence est obligatoire',Ecran.Caption) ;
   GRIDTYP.Col:=ColEtablis ;
   end ;
if (GRIDTYP.Cells[ColDevise,Arow]='')  then
   begin
   PGIBox('La devise est obligatoire',Ecran.Caption) ;
   GRIDTYP.Col:=Coldevise ;
   end ;
if (StrToFloat(GRIDTYP.Cells[ColCoef,Arow])<0) or (StrToFloat(GRIDTYP.Cells[ColCoef,Arow])>100)  then
   begin
   PGIBox('Le coefficient doit être compris entre 0 et 100',Ecran.Caption) ;
   GRIDTYP.Col:=ColCoef ;
   end ;
end ;

function TOF_TarifType.TypeUtiliseDansTarif : boolean ;
var  CodeType : string ;
Q: TQuery ;
begin
  Result:=False;
  CodeType:=GRIDTYP.Cells[ColCode,GRIDTYP.Row];
  Q:=OpenSQL('Select GFM_TARFMODE from TARIFMODE Where GFM_TYPETARIF="'+CodeType+'"',True);
  While not Q.EOF do
  begin
    if ExisteSQL('Select GF_Tarif from Tarif where GF_TARFMODE="'+Q.FindField('GFM_TARFMODE').AsString+'"') then
       begin
       Result:=True ;
       ferme(Q) ;
       Exit ;
       end
       else Q.next ;
  end ;
ferme(Q) ;
end ;

Function TOF_TarifType.EstRempli( Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
if (GRIDTYP.Cells[ColCode,Lig]<>'') then result:= true;
END ;

Function TOF_TarifType.VerifCombo(Champ,Code: String): boolean ;
Var Q: TQuery ; 
begin
if Champ='CODETYP' then
  begin
  Result:= ExisteSQL('Select GFT_CODETYPE From TARIFTYPMODE Where GFT_CODETYPE="'+Code+'"') ;
  end else
  If Champ='DEVISE' then
     begin
     Result:= not ExisteSQL('Select D_DEVISE From DEVISE Where D_DEVISE="'+Code+'"') ;
     end else
     If Champ='ETABLISSEMENT' then
        begin
        Result:= not ExisteSQL('Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="'+Code+'"') ;
        end ;
end ;


{==============================================================================================}
{=============================== Evenement du Grid ========================================}
{==============================================================================================}
procedure TOF_TarifType.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Cancel,Chg: Boolean ;
begin
Cancel := False; Chg := False ;
Case Key of
    VK_F5 : if (Screen.ActiveControl=GRIDTYP) then GSElipsisClick(Sender);
    VK_DOWN : if (Screen.ActiveControl=GRIDTYP) then
            begin
            if NewLigne and EstRempli(GRIDTYP.Row) then OnUpdate
               else if EstRempli(GRIDTYP.Row+1) then GSRowEnter(GRIDTYP,GRIDTYP.Row+1,Cancel,Chg)  //(GRIDTYP.Row:=GRIDTYP.row+1
                  else GRIDTYP.InsertRow(GRIDTYP.Row+1) ;
            end ;
    VK_TAB : If (Screen.ActiveControl=GRIDTYP) then
             if (GRIDTYP.Col=ColCoef) and (NewLigne) and (EstRempli(GRIDTYP.row)) then
                begin
                OnUpdate ;
                if not NewLigne then GSRowEnter(GRIDTYP,GRIDTYP.Row+1,Cancel,Chg) ;
                end ;
    END;

end;

procedure TOF_TarifType.GSElipsisClick(Sender: TObject);
Var DEVISE,ETABLIS: THCritMaskEdit;
    Coord : TRect;
begin
     Inherited ;
if (GRIDTYP.Col = colDevise) then
    BEGIN
    Coord := GRIDTYP.CellRect (GRIDTYP.Col, GRIDTYP.Row);
    DEVISE := THCritMaskEdit.Create (GRIDTYP);
    DEVISE.Parent := GRIDTYP;
    DEVISE.Top := Coord.Top;
    DEVISE.left := Coord.Left;
    DEVISE.Width := 3; DEVISE.Visible := False;
    DEVISE.OpeType:=otString ;
    DEVISE.DATATYPE:='TTDEVISE' ;
    LookUpCombo(DEVISE) ;
    if DEVISE.Text <> '' then GRIDTYP.Cells[GRIDTYP.Col,GRIDTYP.Row]:= DEVISE.Text ;
    DEVISE.Destroy;
    END;
if (GRIDTYP.Col = colEtablis) then
    BEGIN
    Coord := GRIDTYP.CellRect (GRIDTYP.Col, GRIDTYP.Row);
    ETABLIS := THCritMaskEdit.Create (GRIDTYP);
    ETABLIS.Parent := GRIDTYP;
    ETABLIS.Top := Coord.Top;
    ETABLIS.left := Coord.Left;
    ETABLIS.Width := 3; ETABLIS.Visible := False;
    ETABLIS.OpeType:=otString ;
    ETABLIS.DATATYPE:='TTETABLISSEMENT' ;
    LookUpCombo(ETABLIS) ;
    if ETABLIS.Text <> '' then GRIDTYP.Cells[GRIDTYP.Col,GRIDTYP.Row]:= ETABLIS.Text ;
    ETABLIS.Destroy;
    END
END;

procedure TOF_TarifType.GSEnter(Sender: TObject);
Var ACol,ARow : integer;
    Temp : Boolean;
begin
if Action=taConsult then Exit ;
ACol:=GRIDTYP.Col; ARow:=GRIDTYP.Row;
if Not EstRempli(ARow) then NewLigne:=True ;
GRIDTYP.InvalidateRow(ARow);
GSCellEnter(GRIDTYP,ACol,ARow,Temp);
end;

procedure TOF_TarifType.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Action=taConsult then Exit ;
if Not EstRempli(ou) then
begin
     GRIDTYP.DeleteRow(ou) ;
     SetControlEnabled('BInsert',True) ;
     SetControlEnabled('BDelete',True) ;
end ;
GRIDTYP.InvalidateRow(Ou);
end;

procedure TOF_TarifType.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Action=taConsult then Exit ;
if Not EstRempli(Ou) then
    begin
    GRIDTYP.Col:=ColCode ;
    NewLigne:=True;
    end
    else NewLigne:=False ;
GRIDTYP.InvalidateRow(Ou);
end ;

procedure TOF_TarifType.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
If Action=taConsult then Exit ;
GRIDTYP.ElipsisButton := (GRIDTYP.Col = colDevise) or (GRIDTYP.Col = colEtablis) ;
if GRIDTYP.Col=ColCode then
  begin
  if (GRIDTYP.Cells[ColCode,GRIDTYP.Row]<>'') and (not newligne) then
    begin
    GRIDTYP.Col:=colLibelle ;
    end ;
  end;
if GRIDTYP.Col=ColDevise then
  begin
  if (GRIDTYP.Cells[ColDevise,GRIDTYP.Row]<>'') and (not newligne) then
    begin
    GRIDTYP.Col:=colCoef ;
    end ;
  end;
End ;

procedure TOF_TarifType.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
If Action=taConsult then Exit ;
FormateColSaisie(Acol, Arow) ;
If ACol=ColCode then
   if(GRIDTYP.Cells[ColCode,Arow]<>'') and (NewLigne) then
   begin
   GRIDTYP.Cells[ColDevise,Arow]:=V_PGI.DevisePivot ;
   GRIDTYP.Cells[ColCoef,Arow]:='1' ;
   end ;
{If (Acol=ColCode) then
  if(GRIDTYP.Cells[ColCode,Arow]='') then
    begin
    PGIBox('Le code est obligatoire',Ecran.Caption) ;
    GRIDTYP.Col:=ColCode ;
    end ;
If (Acol=ColLibelle) then
   if (GRIDTYP.Cells[ColLibelle,Arow]='')  then
   begin
   PGIBox('Le libelle est obligatoire',Ecran.Caption) ;
   GRIDTYP.Col:=ColLibelle ;
   end ;
If (Acol=ColDevise) then
   if (GRIDTYP.Cells[ColDevise,Arow]='')  then
   begin
   PGIBox('La devise est obligatoire',Ecran.Caption) ;
   GRIDTYP.Col:=Coldevise ;
   end ;}
End ;


{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}
procedure TOF_TarifType.FormateColSaisie (ACol,ARow : Longint ) ;
Var st, Stc:String ;
Begin
St:=GRIDTYP.Cells[ACol,ARow] ; StC:=St ;
if ACol=0 then StC:=Trim(St);
GRIDTYP.Cells[ACol,ARow]:=StC ;
end ;


{==============================================================================================}
{=============================== Mide en forme du grid ========================================}
{==============================================================================================}
procedure TOF_TarifType.ColTriangle ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GRIDTYP.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
  Arect:=GRIDTYP.CellRect(Acol,Arow) ;
  Canvas.Brush.Color := GRIDTYP.FixedColor;
  Canvas.FillRect(ARect);
    if (ARow = GRIDTYP.row) then
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

procedure TOF_TarifType.CodeGras ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);
begin
if (ACol = colCode) and (ARow>0) then
  begin
  Canvas.Font.Style := [fsBold];
  end;
end;


{==============================================================================================}
{=============================== Chargement des types de tarifs ====================================}
{==============================================================================================}
procedure TOF_TarifType.ChargePeriode ;
var Q: TQuery ;
Date,DateDebut,DateFin: String ;
Begin
TobTyp:= TOB.Create('_TARIFTYPMODE',NIL,-1);
Q:=OpenSql('Select * from TARIFTYPMODE ',True) ;
if not Q.EOF then TobTyp.LoadDetailDB('TARIFTYPMODE','','',Q,false,true) ;
If TobTyp.detail.count=0 then Tob.create ('TARIFTYPMODE',TobTyp,-1);
TobTyp.PutGridDetail(GridTyp,True,True,LesColonnes,True);
TobTyp.GetGridDetail(GridTyp,GridTyp.rowcount-1,'TARIFTYPMODE',LesColonnes);
TobTyp.SetAllModifie (false);
ferme(Q) ;
end;


Initialization
registerclasses([TOF_TarifType]);
end.


