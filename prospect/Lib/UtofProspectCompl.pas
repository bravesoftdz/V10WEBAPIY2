unit UtofProspectCompl;

interface
uses  Controls,Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}db,
{$ENDIF}
      forms,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB, vierge,
      SaisUtil,graphics,grids,windows,
      Messages;

Const
   CleChoixCode : String = 'RTC';
   NbTotalChamps : Integer = 60;
   
Type
     TOF_PROSPECTCOMPL = Class (TOF)
     private             // NbTotalChamps
        Oblig : array[1..60] of String;
        Auxiliaire : string ;
        LesColonnes : string ;
        GS : THGRID ;
        //Titres : TStrings;
        ColCpteVte,ColCpteAch : integer ;
        TobToDelete : Tob ;
        NextRang, LastRang : Integer ;
        PasDeSortie : Boolean ;
        ChampTri : String;
        function  GetLastRang : LongInt;
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSDblClick(Sender: TObject);
        procedure GereNewLigne  ;
        procedure NewLigne ;
        Function  EstRempli( Lig : integer) : Boolean ;
        Procedure InitRow (R : integer) ;
        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);

        Function ChercherChamps ( Code : String ) : String;
        procedure MajTitre ( Col : Integer );
        Function ControleZonesOblig(Lig : Integer) : Boolean;
     public
         Action   : TActionFiche ;
         procedure OnArgument (Arguments : String ) ; override ;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         Procedure OnNew  ; override ;
         Procedure OnDelete ; override ;
         Procedure OnCancel ; override ;
         procedure Onclose  ; override ;
     END ;

Function CompteExiste ( Tablette,Valeur : string) : boolean;

const colRang=0 ;
      colAuxi=1 ;
      nbcolfixes=2;

implementation

Procedure TOF_PROSPECTCOMPL.OnArgument (Arguments : String ) ;
var Critere,ChampMul,ValMul : string;
    x,i : integer;
    NbCrit : integer ;
    St,Nam,Valeurs : string ;
    TLib : TOB;
    QQ : TQUERY;
begin
inherited ;
Action:=taModif ;
Repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere<>'' then
    begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='ACTION' then
               begin
               if ValMul='CREATION' then BEGIN Action:=taCreat ; END ;
               if ValMul='MODIFICATION' then BEGIN Action:=taModif ; END ;
               if ValMul='CONSULTATION' then BEGIN Action:=taConsult ; END ;
               end;
           if ChampMul='PROSPECT' then
               Auxiliaire:=ValMul;
        end;
    end;
until  Critere='';

LesColonnes:='RPC_RANG;RPC_AUXILIAIRE' ;

GS:=THGRID(GetControl('G'));
GS.ColCount:=nbcolfixes;

ChampTri:='';
St:='';

TLib:=tob.create('_CHOIXCOD',Nil,-1) ;
// mngcache
QQ := OpenSQL('SELECT CC_CODE,CC_LIBELLE,CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+CleChoixCode+'" and CC_ABREGE<>"" and CC_ABREGE is not Null order by CC_ABREGE', true,-1,'CHAMPSPRO');
TLib.LoadDetailDB('CHOIXCOD','','',QQ,false,true) ;
ferme(QQ);

for i:=0 to TLib.detail.count-1 do
    begin
    Valeurs:=TLib.Detail[i].GetValue('CC_LIBRE');
    if Valeurs <> '' then
      begin
      NbCrit:=0;
      GS.ColCount:=GS.ColCount+1;
      //Titres.Add(TLib.Detail[i].GetValue('CC_LIBELLE'));
      MajTitre(GS.ColCount-1);
      St:= ChercherChamps(TLib.Detail[i].GetValue('CC_CODE'));
      LesColonnes := LesColonnes + ';'+ St;
      Repeat
        Inc(NbCrit);
        Critere:=Trim(ReadTokenSt(Valeurs)) ;
        if NbCrit=1 then
           GS.ColWidths[GS.ColCount-1]:=StrToInt(Critere);
        if NbCrit=2 then
           if Critere = 'X' then
              begin
              ChampTri:=' Order by '+St;
              GS.SortedCol:=GS.ColCount-1;
              end;
        if NbCrit=3 then
           begin
           //GS.ColTypes[GS.ColCount-1]:='B';
           if Critere = 'X' then Oblig[GS.ColCount-1]:='X'
                            else Oblig[GS.ColCount-1]:='-';
           end;
        if NbCrit=4 then
           begin
           if Critere = 'B' then
              begin
              GS.ColTypes[GS.ColCount-1]:='B' ;
              GS.ColFormats[GS.ColCount-1]:=IntToStr(Ord(csCheckbox));
              end
           else if Critere = 'dd/MM/yyyy' then
                GS.ColTypes[GS.ColCount-1]:='D';
           if Critere <> 'B' then
              GS.ColFormats[GS.ColCount-1]:=Critere;
           end;
        if NbCrit=5 then
           begin
           if Critere = 'G' then
              GS.ColAligns[GS.ColCount-1]:=TaLeftJustify
           else if Critere = 'D' then
                begin
                GS.ColTypes[GS.ColCount-1]:='D';
                GS.ColAligns[GS.ColCount-1]:=TaRightJustify;
                end
                else  GS.ColAligns[GS.ColCount-1]:=TaCenter;
           end;
      until  NbCrit=5;
      end;
    end;
    TLib.Free;
if GS.ColCount=nbcolfixes then
   PGIBox('Aucune colonne n''est paramétrée pour cette table','');

GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnRowEnter:=GSRowEnter ;
GS.OnRowExit:=GSRowExit ;
GS.OnRowExit:=GSRowExit ;
GS.OnDblClick:=GSDblClick;

GS.PostDrawCell:= DessineCell;

GS.ColAligns[colRang]:=taCenter;
GS.ColTypes[colRang]:='I' ;

for i:=0 to nbcolfixes-1 do
    GS.ColWidths[i]:=0;

St:=LesColonnes ;

for i:=0 to GS.ColCount-1 do
   BEGIN
   Nam:=ReadTokenSt(St) ;
   if copy(Nam,1,15)='RPC_RPCLIBTABLE' then
      GS.ColFormats[i]:='CB=RTRPCLIBTABLE'+copy(Nam,16,1);
   END ;

AffecteGrid(GS,Action) ;
if GS.ColCount > nbcolfixes then
   TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
if Action=taConsult then GS.Enabled:=False;
end;

Function TOF_PROSPECTCOMPL.ChercherChamps ( Code : String ) : String;
begin
     if copy(Code,1,2) = 'TL' then
        Result := 'RPC_RPCLIBTABLE'+copy(Code,3,1)
     else
       if copy(Code,1,2) = 'TX'  then
          Result := 'RPC_RPCLIBTEXTE'+copy(Code,3,1)
       else
         if copy(Code,1,2) = 'VL' then
            Result := 'RPC_RPCLIBVAL'+copy(Code,3,1)
         else
           if copy(Code,1,2) = 'BL' then
              Result := 'RPC_RPCLIBBOOL'+copy(Code,3,1)
         else
              Result := 'RPC_RPCLIBDATE'+copy(Code,3,1)
         ;
end;

Procedure TOF_PROSPECTCOMPL.OnLoad  ;
var QQ : TQUERY;
begin
inherited ;
LastRang:=GetLastRang;
NextRang:=LastRang+1;

LaTob:=tob.create('_PROSPECTCOMPL',Nil,-1) ;
//Order by ChampTri
QQ:=OpenSQL('SELECT * from PROSPECTCOMPL where RPC_AUXILIAIRE="'+Auxiliaire+'"'+ChampTri,false);
Latob.LoadDetailDB('PROSPECTCOMPL','','',QQ,false,true) ;
ferme(QQ);
If LaTob.detail.count=0 then
   begin
   Tob.create ('PROSPECTCOMPL',LaTob,-1) ;
   LaTob.detail[0].InitValeurs(false);
   LaTob.detail[0].PutValue('RPC_AUXILIAIRE',Auxiliaire);
   Latob.PutGridDetail(GS,True,True,LesColonnes,True);
   InitRow( GS.RowCount-1 ) ;
   end
else
   Latob.PutGridDetail(GS,True,True,LesColonnes,True);
NewLigne ;
TobToDelete:=tob.create('_PROSPECTCOMPL à supprimer',Nil,-1) ;
SetControlEnabled('BVENTILACH',(GS.cells[ColCpteAch,GS.row]<>'')) ;
SetControlEnabled('BVENTILVTE',(GS.cells[ColCpteVte,GS.row]<>'')) ;

end;

Procedure TOF_PROSPECTCOMPL.OnUpdate  ;
var  i : integer ;
begin
inherited ;
for i := 1 to GS.RowCount-1 do
  if EstRempli(i) and ControleZonesOblig(i) then
     begin
     PasDeSortie := true ;
     break;
     end;

if PasDeSortie then exit;

application.processMessages;

if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then
        GS.ShowCombo(GS.Col,GS.Row) ;


For i:=GS.fixedrows to GS.RowCount-2 do
   begin
   if GS.cells[colRang,i]='' then begin GS.cells[colRang,i]:=inttostr(NextRang); inc(NextRang); end ;
   end ;
if GS.rowcount > 2 then
    begin
    LaTob.GetGridDetail( THGRID(GetControl('G')),GS.rowcount-2,'PROSPECTCOMPL',LesColonnes) ;
    //LaTob.PutValue('RPC_AUXILIAIRE',Auxiliaire);
    LaTob.InsertOrUpdateDB(True) ;
    LaTob.SetAllModifie (false);
    end ;

TobToDelete.DeleteDB (true) ;
TobToDelete.ClearDetail ;

end;

Procedure TOF_PROSPECTCOMPL.OnClose ;
var i : integer ;
begin
inherited ;
if PasDeSortie then
   begin
   LastError := 1 ;
   PasDeSortie:=False;
   exit;
   end;

//Titres.Free;
PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
application.processMessages;
LastError:=0;
if (GS.ColCount > nbcolfixes) and (GS.rowcount > 2) then
      LaTob.GetGridDetail( THGRID(GetControl('G')),GS.rowcount-2,'PROSPECTCOMPL',LesColonnes) ;
for i:=0 to LaTob.detail.count-1 do
    begin
    if (LaTob.Detail[i].IsOneModifie) or (TobToDelete.detail.count > 0) then
   //if LaTob.IsOneModifie then
        begin
        if PGIAsk('Voulez-vous enregistrer les modifications ?',Ecran.Caption)=mrYes then
            begin
            LastError:=0;
            OnUpdate;
            if LastError<>0 then exit;
            break;
            end else
            begin
            LastError:=0;
            break;
            end;
        end;
    end;
LaTob.free ; Latob:=nil;
TobToDelete.free ;
end;


Procedure TOF_PROSPECTCOMPL.OnCancel ;
var i : integer ;
begin
inherited ;
For i:=GS.fixedrows to GS.RowCount-2 do
   begin
   if (GS.cells[colRang,i]<>'') and (strtoint(GS.cells[colRang,i])> LastRang) then
        begin
       // ExecuteSQL('DELETE FROM VENTIL WHERE (V_NATURE like "HA%" OR V_NATURE like "HV%") and '
       //                  +' V_COMPTE="'+GS.cells[colRang,i]+'"') ;
        end;
   end ;
TobToDelete.free ;  TobToDelete:=Nil ;
LaTob.free ; LaTob:=Nil ;
OnLoad ;
end ;

Procedure TOF_PROSPECTCOMPL.OnNew  ;
begin
inherited ;
GS.InsertRow(GS.row) ;
GS.row:=gs.row-1 ;
if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then
       GS.ShowCombo(GS.Col,GS.Row) ;
InitRow( GS.Row ) ;
end;

Procedure TOF_PROSPECTCOMPL.InitRow (R : integer) ;
var i : integer ;
begin
for i:=0 to GS.ColCount do
   begin
   if GS.ColTypes[i]='D' then GS.cells[i,R]:=DateToStr(iDate1900)
   //else if GS.ColTypes[i]='I' then GS.cells[i,R]:='0'
   else GS.cells[i,R]:='';
   end;
GS.Cells[ColAuxi,R]:=Auxiliaire;
end ;

Function TOF_PROSPECTCOMPL.GetLastRang : LongInt;
var Q : TQuery;
begin
Q := OpenSQL('SELECT MAX(RPC_RANG) FROM PROSPECTCOMPL',true);
if not Q.EOF then result := Q.Fields[0].AsInteger
             else result := 0;
Ferme(Q);
end;

Procedure TOF_PROSPECTCOMPL.OnDelete  ;
begin
inherited ;
if GS.Row<=0 then Exit ;
//if GS.RowCount<=2 then Exit ;
if EstRempli(GS.Row) then
    begin
    if GS.objects[0,GS.row]<>Nil then TOB(GS.objects[0,GS.row]).ChangeParent( TobToDelete,-1) ;
    GS.DeleteRow(GS.row) ;
    end;
end;

procedure TOF_PROSPECTCOMPL.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
GereNewLigne;
if Not Cancel then
   BEGIN
   GS.ElipsisButton:=((GS.Col=ColCpteAch) or (GS.Col=ColCpteVte)) ;
   end ;
end;

procedure TOF_PROSPECTCOMPL.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
end;


procedure TOF_PROSPECTCOMPL.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_PROSPECTCOMPL.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Not EstRempli(ou) then GS.DeleteRow(ou)
else
    ControleZonesOblig(ou);
GS.InvalidateRow(ou) ;
end;

Function TOF_PROSPECTCOMPL.ControleZonesOblig(Lig : Integer) : Boolean;
var i: Integer;
begin
    Result:=false;
    for i:=1 to GS.ColCount-1 do
       begin
       if (GS.ColTypes[i]='D') and (GS.cells[i,Lig]=DateToStr(iDate1900)) and ( Oblig[i]='X' ) then
           begin
           Result:= true;
           break;
           end
       else
           if (GS.ColTypes[i]<>'D') and (GS.Cells[i,Lig]='') and ( Oblig[i]='X' )then
             begin
             Result:= true;
             break;
             end;
       end;
    If Result then
       begin
       PGIBox('Cette zone obligatoire n''est pas renseignée','');
       GS.Row:=Lig;
       GS.Col:=i;
       end;
end;

procedure TOF_PROSPECTCOMPL.GSDblClick (Sender: TObject);
begin
if GS.ColTypes[GS.col]='B' then
   if GS.cells[GS.col,Gs.Row]='X' then
   GS.cells[GS.col,Gs.Row]:='-'
   else
   GS.cells[GS.col,Gs.Row]:='X';

end;

procedure TOF_PROSPECTCOMPL.GereNewLigne  ;
BEGIN
   if EstRempli(GS.RowCount-1) then NewLigne else
      if Not EstRempli(GS.RowCount-2) then GS.RowCount:=GS.RowCount-1 ;
END ;

procedure TOF_PROSPECTCOMPL.NewLigne ;
BEGIN
    GS.RowCount:=GS.RowCount+1 ;
    InitRow( GS.RowCount-1 ) ;
END ;

Function TOF_PROSPECTCOMPL.EstRempli( Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
for i:=nbcolfixes to GS.ColCount-1 do
  begin
   if ((GS.ColTypes[i]='D') and (GS.cells[i,Lig]<>DateToStr(iDate1900)) ) then
       begin
       result:= true;
       break;
       end
   else
       if (GS.ColTypes[i]<>'D') and (GS.Cells[i,Lig]<>'') then
         begin
         result:= true;
         break;
         end;
  end;
END ;

procedure TOF_PROSPECTCOMPL.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

Function CompteExiste ( Tablette,Valeur : string) : boolean;
var St : string;
BEGIN
Result := False;
St := RechDom (Tablette, Valeur, False);
if (St <> '') AND (St <> 'Error') then Result := True;
END;

procedure TOF_PROSPECTCOMPL.MajTitre ( Col : Integer );
var F : array[1..62] of string;
    A : array[1..62] of TAlignment;
    T : array[1..62] of Char;
    i : integer;
begin
for i:=1 to GS.ColCount-1 do begin
F[i]:= GS.ColFormats[i];
A[i]:= GS.ColAligns[i];
T[i]:= GS.ColTypes[i];
end;
//GS.Titres:=Titres;
for i:=1 to GS.ColCount-1 do begin
GS.ColFormats[i]:= F[i];
GS.ColAligns[i]:= A[i];
GS.ColTypes[i]:= T[i];
end;
end;

Initialization
registerclasses([TOF_PROSPECTCOMPL]);
end.
