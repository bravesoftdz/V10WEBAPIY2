unit UtofCodeCpta;

interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      forms,sysutils,ComCtrls,grids,windows,
      HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics,Ventil,
      M3FP,HTB97,Messages,UTofOptionEdit, ParamSoc ;

Var PEtat : RPARETAT;

Type
     TOF_CODECPTA = Class (TOF)
     private
        LesColonnes : string ;
        GS : THGRID ;
        ColCptaArt,ColCpteVte,ColCpteAch,ColCpteStock,ColCpteVarStk,ColTiers,ColVenteAchat,ColEtabliss : integer ;
        TobToDelete : Tob ;
        NextRang, LastRang : Integer ;
        NoAchat,AvecImmoDiv,AvecStock : Boolean ;
        BImprimer : TToolBarButton97;
        ValeurCell : string;
        function  GetLastRang : LongInt;
        procedure GSElipsisClick(Sender: TObject);
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GereNewLigne  ;
        procedure NewLigne ;
        Function  EstRempli( Lig : integer) : Boolean ;
        Procedure InitRow (R : integer) ;
        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;AState: TGridDrawState);
    		procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
        procedure VentilCompteHT(Nature : String) ;
        Function  VerifTout : Boolean ;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure BImprimerClick(Sender: TObject);
        procedure PrepareGrid;
        function  PrepareImpression : integer ;
     protected
        procedure Onclose  ; override ;
     public
         Action   : TActionFiche ;
         procedure OnArgument (Arguments : String ) ; override ;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         Procedure OnNew  ; override ;
         Procedure OnDelete ; override ;
         Procedure OnCancel ; override ;

     END ;

Function CompteExiste ( Tablette,Valeur : string) : boolean;

const colRang=1 ;

implementation

Procedure TOF_CODECPTA.OnArgument (Arguments : String ) ;
var i,NbCol : integer ;
    St : string ;
    //uniquement en line
    //BT : TToolbarButton97;

begin
inherited ;
AvecImmoDiv:=False ; AvecStock:=False ;
{$IFNDEF CCS3}
AvecImmoDiv:=GetParamSoc('SO_GCCPTAIMMODIV') ;
AvecStock:=GetParamSoc('SO_GCINVPERM') ;
{$ENDIF}

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
St:=Arguments ;
NoAchat:=False ;
i:=Pos('NOACHAT',St) ;
if i>0 then  NoAchat:=True ;
//NoAchat:=True ;
NbCol:=8;
GS:=THGRID(GetControl('G'));
GS.OnElipsisClick:=GSElipsisClick  ;
GS.OnDblClick:=GSElipsisClick ;
GS.OnCellEnter:=GSCellEnter ;
GS.OnCellExit:=GSCellExit ;
GS.OnRowEnter:=GSRowEnter ;
GS.OnRowExit:=GSRowExit ;
GS.PostDrawCell:= DessineCell;
GS.GetCellCanvas := GetCellCanvas;
GS.ColCount:=NbCol;
GS.ColAligns[colRang]:=taCenter;
GS.ColTypes[colRang]:='I' ;
GS.ColWidths[1]:=0;   // rang
GS.ColWidths[0]:=15;
PrepareGrid;
AffecteGrid(GS,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
BImprimer:=TToolBarButton97(GetControl('BIMPRIMER'));
BImprimer.OnClick:=BImprimerClick;
PEtat.Tip:='E'; PEtat.Nat:='PAR'; PEtat.Modele:='VCP'; PEtat.Titre:=TFVierge(Ecran).Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL:='';
GS.Col := 2; GS.Row := 1;
GS.SetFocus;
//uniquement en line
{*
  BT := TToolbarButton97 (GetCOntrol('BVENTILVTE'));
  if BT <> nil then BT.visible := false;
*}
end;

procedure TOF_CODECPTA.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_F5 : if (Screen.ActiveControl=GS) then GSElipsisClick(Sender);
    END;
end;

procedure TOF_CODECPTA.BImprimerClick(Sender: TObject);
begin
EntreeOptionEditNP(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 Nil,PEtat.stSQL,PEtat.Titre,PrepareImpression);
end;

Procedure TOF_CODECPTA.OnLoad  ;
begin
inherited ;
LastRang:=GetLastRang;
NextRang:=LastRang+1;

LaTob:=tob.create('_CODECPTA',Nil,-1) ;
Latob.LoadDetailDB('CODECPTA','','',nil,false,true) ;
If LaTob.detail.count=0 then Tob.create ('CODECPTA',LaTob,-1) ;
Latob.PutGridDetail(GS,True,True,LesColonnes,True);
GS.RowCount:=GS.RowCount+1 ;
ValeurCell := GS.Cells[2, 1];
TobToDelete:=tob.create('_CODECPTA à supprimer',Nil,-1) ;
SetControlEnabled('BVENTILACH',(GS.cells[ColCpteAch,GS.row]<>'')) ;
SetControlEnabled('BVENTILVTE',(GS.cells[ColCpteVte,GS.row]<>'')) ;
if AvecStock then SetControlEnabled('BVENTILSTK',(GS.cells[ColCpteVarStk,GS.row]<>'')) ;

GS.Col:=2 ; GS.Row:=1 ;
if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then GS.ShowCombo(GS.Col,GS.Row) ;
//uniquement en line
//GS.cells[ColTiers,0] := 'Famille Comptable client'; mis en commentaire par BRL 310309
//GS.ColWidths [ColTiers] := 120;

GS.SetFocus ; // mettre ici sinon combo ligne 1 écrasé par valeur du dernier combo affiché
end;

Procedure TOF_CODECPTA.OnUpdate  ;
var  i : integer ;
begin
inherited ;
if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then PostMessage(GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB,  0) ;
Application.processMessages;

//if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then GS.ShowCombo(GS.Col,GS.Row) ;

if not VerifTout then  Exit;
For i:=GS.fixedrows to GS.RowCount-2 do
   begin
   if GS.cells[colRang,i]='' then begin GS.cells[colRang,i]:=inttostr(NextRang); inc(NextRang); end ;
   end ;
if GS.rowcount > 2 then
    begin
    LaTob.GetGridDetail( THGRID(GetControl('G')),GS.rowcount-2,'CODECPTA',LesColonnes) ;
    LaTob.InsertOrUpdateDB(True) ;
    LaTob.SetAllModifie (false);
    end ;
// suppressions
if TobToDelete.detail.count > 0 then
    for i:=0 to TobToDelete.detail.count-1 do
       ExecuteSQL('DELETE FROM VENTIL WHERE (V_NATURE like "HA%" OR V_NATURE like "HV%" OR V_NATURE like "ST%") and '
                 +' V_COMPTE="'+string(TobToDelete.Detail[i].GetValue('GCP_RANG'))+'"') ;
TobToDelete.DeleteDB (true) ;
TobToDelete.ClearDetail ;

if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then GS.ShowCombo(GS.Col,GS.Row) ;

end;

Procedure TOF_CODECPTA.OnClose ;
var i : integer ;
begin
inherited ;
PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
application.processMessages;
LastError:=0;
for i:=0 to LaTob.detail.count-1 do
    begin
    if (LaTob.Detail[i].IsOneModifie) or (TobToDelete.detail.count > 0) then
        begin
        if PGIAsk('Voulez-vous enregistrer les modifications ?',Ecran.Caption)=mrYes then
           begin
           LastError:=0; OnUpdate;
           if LastError<>0 then exit;
           break;
           end else
           begin
           LastError:=0; break;
           end;
        end;
    end;
LaTob.free ; Latob:=nil;
TobToDelete.free ;
end;

Function TOF_CODECPTA.VerifTout : Boolean ;
var  i : integer ;
begin
Result:= True ;
For i:=GS.fixedrows to GS.RowCount-2 do
   begin
    If (GS.cells[ColCpteAch,i]='') and (GS.cells[ColCpteVte,i]='') then
        begin
        LastError:=1; LastErrorMsg:='Vous devez renseigner un compte comptable' ;
        Result:=false;
        GS.row:=i; if NoAchat then GS.Col:=ColCpteVte else GS.Col:=ColCpteAch;
        break ;
        end ;
    If (Not NoAchat) and (GS.cells[ColCpteAch,i]<>'') and
       (not CompteExiste('TZGCHARGE',GS.cells[ColCpteAch,i])) then
       BEGIN
       if (Not AvecImmoDiv) or
          ((Not CompteExiste('TZGIMMO',GS.cells[ColCpteAch,i])) and
           (Not CompteExiste('TZGDIVERS',GS.cells[ColCpteAch,i])))
          then
           begin
           LastError:=2; LastErrorMsg:='Ce compte n''existe pas' ;
           Result:=false;
           GS.row:=i; GS.Col:=ColCpteAch;
           break ;
           end ;
        END ;
   If  (GS.cells[ColCpteVte,i]<>'') and (not CompteExiste('TZGPRODUIT',GS.cells[ColCpteVte,i])) then
        begin
        if (Not AvecImmoDiv) or
           ((Not CompteExiste('TZGIMMO',GS.cells[ColCpteVte,i])) and
            (Not CompteExiste('TZGDIVERS',GS.cells[ColCpteVte,i]))) then
            BEGIN
            LastError:=2; LastErrorMsg:='Ce compte n''existe pas' ;
            Result:=false;
            GS.row:=i; GS.Col:=ColCpteVte;
            break ;
            end;
        end ;
   if AvecStock then
      BEGIN
      if  (GS.cells[ColCpteStock,i]<>'') and (not CompteExiste('TZGDIVERS',GS.cells[ColCpteStock,i])) then
           BEGIN
           LastError:=2; LastErrorMsg:='Ce compte n''existe pas' ;
           Result:=False ; GS.row:=i; GS.Col:=ColCpteStock ;
           break ;
           END ;
      if  (GS.cells[ColCpteVarStk,i]<>'') and
            (not CompteExiste('TZGCHARGE',GS.cells[ColCpteVarStk,i])) and
            (not CompteExiste('TZGPRODUIT',GS.cells[ColCpteVarStk,i])) then
           BEGIN
           if (Not AvecImmoDiv) or
             ((Not CompteExiste('TZGIMMO',GS.cells[ColCpteVarStk,i])) and
              (Not CompteExiste('TZGDIVERS',GS.cells[ColCpteVarStk,i]))) then
              BEGIN
              LastError:=2; LastErrorMsg:='Ce compte n''existe pas' ;
              Result:=False ; GS.row:=i; GS.Col:=ColCpteVarStk ;
              break ;
              END ;
           END ;
      END ;
   end ;
end ;


Procedure TOF_CODECPTA.OnCancel ;
var i : integer ;
begin
inherited ;
For i:=GS.fixedrows to GS.RowCount-2 do
   begin
   if (GS.cells[colRang,i]<>'') and (strtoint(GS.cells[colRang,i])> LastRang) then
        begin
        ExecuteSQL('DELETE FROM VENTIL WHERE (V_NATURE like "HA%" OR V_NATURE like "HV%" OR V_NATURE like "ST%") and '
                  +' V_COMPTE="'+GS.cells[colRang,i]+'"') ;
        end;
   end ;
TobToDelete.free ;  TobToDelete:=Nil ;
LaTob.free ; LaTob:=Nil ;
OnLoad ;
end ;

Procedure TOF_CODECPTA.OnNew  ;
begin
inherited ;
GS.InsertRow(GS.row) ; GS.row:=gs.row-1 ;
if Copy(GS.ColFormats[GS.Col],1,3)='CB=' then GS.ShowCombo(GS.Col,GS.Row) ;
end;

Procedure TOF_CODECPTA.InitRow (R : integer) ;
var i : integer ;
begin
for i:=0 to GS.ColCount do GS.cells[i,R]:='';
end ;

Function TOF_CODECPTA.GetLastRang : LongInt;
var Q : TQuery;
begin
Q := OpenSQL('SELECT MAX(GCP_RANG) FROM CODECPTA',true);
if not Q.EOF then result := Q.Fields[0].AsInteger
             else result := 0;
Ferme(Q);
end;

Procedure TOF_CODECPTA.OnDelete  ;
begin
inherited ;
if GS.Row<=0 then Exit ;
if GS.RowCount<=2 then Exit ;
if GS.objects[0,GS.row]<>Nil then TOB(GS.objects[0,GS.row]).ChangeParent( TobToDelete,-1) ;
GS.DeleteRow(GS.row) ;
end;

procedure TOF_CODECPTA.GSElipsisClick(Sender: TObject);
var criteres : string ;
begin
if GS.Col=ColCpteVte then
   begin
   if AvecImmoDiv then Criteres:='(G_NATUREGENE="PRO" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
                  else Criteres:='G_NATUREGENE="PRO"' ;
   LookupList(GS,TraduireMemoire('Comptes de produits'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
   end ;
if GS.Col=ColCpteAch then
   begin
   if AvecImmoDiv then Criteres:='(G_NATUREGENE="CHA" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
                  else Criteres:='G_NATUREGENE="CHA"' ;
   LookupList(GS,TraduireMemoire('Comptes de charges'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
   end ;
if ((AvecStock) and (GS.Col=ColCpteStock)) then
   begin
   Criteres:='G_NATUREGENE="DIV"' ;
   LookupList(GS,TraduireMemoire('Comptes divers de stock'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
   end ;
if ((AvecStock) and (GS.Col=ColCpteVarStk)) then
   begin
   if AvecImmoDiv then Criteres:='(G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" OR G_NATUREGENE="IMO" OR G_NATUREGENE="DIV")'
                  else Criteres:='G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" ' ;
   LookupList(GS,TraduireMemoire('Comptes de variation de stock'),'GENERAUX','G_GENERAL','G_LIBELLE',Criteres,'G_GENERAL',TRUE, 0) ;
   end ;
end;

procedure TOF_CODECPTA.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action=taConsult then Exit ;
  if Not Cancel then
   BEGIN
  	GereNewLigne;
   GS.ElipsisButton:=((GS.Col=ColCpteAch)   or (GS.Col=ColCpteVte) or
                      (GS.Col=ColCpteStock) or (GS.Col=ColCpteVarStk)) ;
    ValeurCell:=GS.Cells[GS.Col, GS.Row] ;
   end ;
end;

procedure TOF_CODECPTA.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var ZZ : WideString;
begin
  if Action=taConsult then Exit ;
  ZZ := GS.CellValues[ColventeAchat,Arow];
  if (ACol=ColCpteAch) then  SetControlEnabled('BVENTILACH',(GS.cells[ColCpteAch,ARow]<>'')) ;
  if (ACol=ColCpteVte) then  SetControlEnabled('BVENTILVTE',(GS.cells[ColCpteVte,ARow]<>'')) ;
  if ((ACol=ColCpteVarStk) and (AvecStock)) then SetControlEnabled('BVENTILSTK',(GS.cells[ColCpteVarStk,ARow]<>'')) ;
  //
  if (Acol <> ColCptaArt) and (ACol <> ColCpteAch) and (ACol <> ColVenteAchat) and (ACol <> ColEtabliss) and (ZZ='CON') then
  begin
    GS.Cells[ACol,Arow]:='';
  end;
  //
  if (ValeurCell <> GS.Cells[ACol, ARow]) and (TOB(GS.Objects[0, ARow])<> Nil) then
    TOB(GS.Objects[0, ARow]).SetAllModifie(True);
end;

procedure TOF_CODECPTA.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
SetControlEnabled('BVENTILACH',(GS.cells[ColCpteAch,ou]<>'')) ;
SetControlEnabled('BVENTILVTE',(GS.cells[ColCpteVte,ou]<>'')) ;
if AvecStock then SetControlEnabled('BVENTILSTK',(GS.cells[ColCpteVarStk,ou]<>'')) ;
end;

procedure TOF_CODECPTA.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Not EstRempli(ou) then GS.DeleteRow(ou) ;
GS.InvalidateRow(ou) ;
end;

procedure TOF_CODECPTA.GereNewLigne  ;
BEGIN
if EstRempli(GS.RowCount-1) then NewLigne else
   if Not EstRempli(GS.RowCount-2) then GS.RowCount:=GS.RowCount-1 ;
END ;

procedure TOF_CODECPTA.NewLigne ;
BEGIN
GS.RowCount:=GS.RowCount+1 ;
InitRow( GS.RowCount-1 ) ;
END ;

Function TOF_CODECPTA.EstRempli( Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
for i:=1 to GS.ColCount-1 do
  if (GS.Cells[i,Lig]<>'') then begin result:= true; break; end;
END ;

procedure TOF_CODECPTA.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
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

procedure AGLVentilCompteHT( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  if (ToTof is TOF_CODECPTA) then TOF_CODECPTA(ToTof).VentilCompteHT(string(Parms[1])) else exit;
end;

procedure TOF_CODECPTA.VentilCompteHT(Nature : String) ;
Var StAxes : String ;
begin
if GS.cells[colRang,GS.row]='' then begin GS.cells[colRang,GS.row]:=inttostr(NextRang); inc(NextRang); end ;
if EstSerie(S7) then StAxes:='12345' else
   BEGIN
   StAxes:='1' ;
   {$IFNDEF CCS3}
   if VH_GC.GCventAxe2 then StAxes:=StAxes+'2' ;
   if VH_GC.GCventAxe3 then StAxes:=StAxes+'3' ;
  {$ENDIF}
   END ;
ParamVentil(Nature,string(GS.cells[colRang,GS.row]) ,StAxes,taCreat,FALSE);
end ;

Function CompteExiste ( Tablette,Valeur : string) : boolean;
var St : string;
BEGIN
Result := False;
St := RechDom (Tablette, Valeur, False);
if (St <> '') AND (St <> 'Error') then Result := True;
END;

procedure TOF_CODECPTA.PrepareGrid;
var i,NbCol : integer ;
    St,Nam : string ;
begin
LesColonnes:='FIXED;GCP_RANG;GCP_VENTEACHAT' ;
NbCol := 8;   // PA
if GetParamSocSecur('SO_GCVENTCPTAART', False) then begin LesColonnes:=LesColonnes+';GCP_COMPTAARTICLE'; inc(NbCol); end ;
if GetParamSocSecur('SO_GCVENTCPTATIERS', False) then begin LesColonnes:=LesColonnes+';GCP_COMPTATIERS'; inc(NbCol); end ;
if GetParamSocSecur('SO_GCVENTCPTAAFF', False) then begin LesColonnes:=LesColonnes+';GCP_COMPTAAFFAIRE'; inc(NbCol); end ;
LesColonnes :=LesColonnes+';GCP_ETABLISSEMENT;GCP_REGIMETAXE;GCP_FAMILLETAXE;GCP_CPTEGENEACH;GCP_CPTEGENEVTE' ;
if AvecStock then begin LesColonnes:=LesColonnes+';GCP_CPTEGENESTOCK;GCP_CPTEGENEVARSTK'; inc(NbCol,2); end ;
St:=LesColonnes ;
GS.ColCount:=NbCol; // PA
for i:=0 to GS.ColCount-1 do
   BEGIN
   if i>1 then  GS.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='GCP_COMPTAARTICLE' then
   begin
		ColCptaArt := i;
   	GS.ColFormats[i]:='CB=GCCOMPTAARTICLE||<<Tous>>';
		GS.ColWidths [I] := 120;
   end else if Nam='GCP_COMPTATIERS' then
   begin
   	GS.ColFormats[i]:='CB=GCCOMPTATIERS||<<Tous>>';
    ColTiers := i;
   end else if Nam='GCP_VENTEACHAT' then
   begin
   	GS.ColFormats[i]:='CB=GCVENTEACHAT|AND (CO_CODE="ACH" OR CO_CODE="VEN" OR CO_CODE="CON")|<<Toute>>';
    ColVenteAchat := i;
   end else if Nam='GCP_COMPTAAFFAIRE' then GS.ColFormats[i]:='CB=AFCOMPTAAFFAIRE||<<Tous>>'
   else if Nam='GCP_ETABLISSEMENT' then
   begin
   	GS.ColFormats[i]:='CB=TTETABLISSEMENT||<<Tous>>';
    ColEtabliss := i;
   end else if Nam='GCP_REGIMETAXE' then GS.ColFormats[i]:='CB=TTREGIMETVA||<<Tous>>'
   else if Nam='GCP_FAMILLETAXE' then BEGIN GS.ColFormats[i]:='CB=GCFAMILLETAXE1||<<Tous>>'; GS.ColWidths[i]:=120; END
   else if Nam='GCP_CPTEGENEACH' then ColCpteAch:=i
   else if Nam='GCP_CPTEGENEVTE' then ColCpteVte:=i
   else if Nam='GCP_CPTEGENESTOCK' then ColCpteStock:=i
   else if Nam='GCP_CPTEGENEVARSTK' then ColCpteVarStk:=i
   ;
   END ;
if NoAchat then
   begin
   GS.ColWidths[ColCpteAch]:=0;   // Masque Compte d'achat
   SetControlVisible('BVENTILACH',False) ;
   end ;
if Not AvecStock then SetControlVisible('BVENTILSTK',False) ;
end;

function TOF_CODECPTA.PrepareImpression : integer ;
begin
Result:=0;
end;

procedure TOF_CODECPTA.GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  //
end;

Initialization
registerclasses([TOF_CODECPTA]);
RegisterAglProc( 'VentilCompteHT', TRUE , 1, AGLVentilCompteHT );
end.
