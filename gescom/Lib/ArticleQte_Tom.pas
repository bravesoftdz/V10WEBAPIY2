{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/02/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ARTICLEQTE 
Mots clefs ... : TOM;GCARTICLEQTE
*****************************************************************}
Unit ArticleQte_Tom;

Interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     eFiche,eFichList,maineagl,
{$ELSE}
     db,dbtables,Fiche,FichList,Fe_main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTB97,SaisUtil,Vierge,
     UTOM,UTob,UTilFonctionCalcul ;

Type
  TOM_ARTICLEQTE = Class (TOM)
    Private
    BFORMULEQTEVTE : TToolbarButton97 ;
    BFORMULEQTEACH : TToolbarButton97 ;
    BVTEACH : TToolbarButton97 ;
    BVTESUP : TToolbarButton97 ;
    BACHVTE : TToolbarButton97 ;
    BACHSUP : TToolbarButton97 ;
    GVte : THGrid ;
    GAch : THGrid ;
    VLib,VVal,VEna,VVis : integer ;
    ALib,AVal,AEna,AVis : integer ;
    LesColVte,LesColAch : string ;
    TobVarV,TobVarA : Tob;
    procedure InsertionBaseArtFormule(TobRef : TOB) ;
    procedure EtudieColsListe;
    procedure ChargeGrille ;
    procedure RempliTob (Flux : string; RFonct : R_FonctCal);
    Function  GetTOBLigne (Flux : string; ARow : integer) : TOB ;
    procedure TraiterValeur(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TraiterEnable(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TraiterVisible(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure SortDuGrid;
    // Evènements du Grid
    procedure GVarExit(Sender: TObject);
    procedure GVarCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GVarCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GVarDblClick(Sender: TObject);
    procedure GVarRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GVarRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

    procedure AppelFormuleQte(Sender : TObject);
    procedure ActivationBoutonFormule ;
    procedure CopieFormuleClick(Sender : TObject);
    procedure DelFormuleClick(Sender : TObject);
    Public
    Action   : TActionFiche ;
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( StArg: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}        'Le code article doit être renseigné'
          {2}       ,''
                     );

Implementation

{==============================================================================================}
{============================ Procédure de la TOF METHODEAPPRO ================================}
{==============================================================================================}
procedure TOM_ARTICLEQTE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_ARTICLEQTE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ARTICLEQTE.OnUpdateRecord ;
begin
  Inherited ;
if GetField('GAF_CODEARTICLE')='' then
  begin
  LastError:=1;
  LastErrorMsg:=TexteMessage[LastError];
  exit
  end;
end ;

procedure TOM_ARTICLEQTE.OnAfterUpdateRecord ;
begin
  Inherited ;
ExecuteSql('Delete ARTFORMULEVAR where GAV_ARTICLE="'+
            GetField('GAF_ARTICLE')+'"');
InsertionBaseArtFormule(TobVarV);
InsertionBaseArtFormule(TobVarA);
end ;

procedure TOM_ARTICLEQTE.OnLoadRecord ;
var QQ : TQuery;
begin
  Inherited ;
if GetField('GAF_CODEARTICLE')='' then SetField('GAF_CODEARTICLE',Copy(GetField('GAF_ARTICLE'),1,18)) ;
ActivationBoutonFormule ;
QQ:=OpenSql('Select * From ARTFORMULEVAR where GAV_VENTEACHAT="VTE" and GAV_ARTICLE="'+
            GetField('GAF_ARTICLE')+'"',True);
TobVarV.LoadDetailDB('ARTFORMULEVAR','','',QQ,False);
Ferme(QQ);
QQ:=OpenSql('Select * From ARTFORMULEVAR where GAV_VENTEACHAT="ACH" and GAV_ARTICLE="'+
            GetField('GAF_ARTICLE')+'"',True);
TobVarA.LoadDetailDB('ARTFORMULEVAR','','',QQ,False);
Ferme(QQ);
ChargeGrille;
end ;

procedure TOM_ARTICLEQTE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_ARTICLEQTE.OnArgument ( StArg: String ) ;
Var St : string ;
    i : integer ;
begin
  Inherited ;
St:=StArg ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;

BFORMULEQTEVTE:=TToolbarButton97(GetControl('BFORMULEQTEVTE')) ;
BFORMULEQTEACH:=TToolbarButton97(GetControl('BFORMULEQTEACH')) ;
BFORMULEQTEVTE.OnClick:=AppelFormuleQte;
BFORMULEQTEACH.OnClick:=AppelFormuleQte;
BVTEACH:=TToolbarButton97(GetControl('BVTEACH')) ;
BVTESUP:=TToolbarButton97(GetControl('BVTESUP')) ;
BACHVTE:=TToolbarButton97(GetControl('BACHVTE')) ;
BACHSUP:=TToolbarButton97(GetControl('BACHSUP')) ;
BVTEACH.OnClick:=CopieFormuleClick;
BACHVTE.OnClick:=CopieFormuleClick;
BVTESUP.OnClick:=DelFormuleClick;
BACHSUP.OnClick:=DelFormuleClick;

GVte:=THGrid(GetControl('GVTE'));
GVte.ListeParam:='GCARTFORMULEVAR';
GVte.OnExit:=GVarExit;
GVte.OnCellEnter:=GVarCellEnter ;
GVte.OnCellExit:=GVarCellExit ;
GVte.OnRowEnter:=GVarRowEnter ;
GVte.OnRowExit:=GVarRowExit ;
GVte.OnDblClick:=GVarDblClick ;

GAch:=THGrid(GetControl('GACH'));
GAch.ListeParam:='GCARTFORMULEVAR';
GAch.OnExit:=GVarExit;
GAch.OnCellEnter:=GVarCellEnter ;
GAch.OnCellExit:=GVarCellExit ;
GAch.OnRowEnter:=GVarRowEnter ;
GAch.OnRowExit:=GVarRowExit ;
GAch.OnDblClick:=GVarDblClick ;

if TFFiche(Ecran).TypeAction=taConsult then
  begin
  SetControlEnabled('BFORMULEQTEVTE',False) ;
  SetControlEnabled('BFORMULEQTEACH',False) ;
  SetControlEnabled('BVTEACH',False) ;
  SetControlEnabled('BVTESUP',False) ;
  SetControlEnabled('BACHVTE',False) ;
  SetControlEnabled('BACHSUP',False) ;
  end;
TobVarV:=Tob.Create('',nil,-1);
TobVarA:=Tob.Create('',nil,-1);
EtudieColsListe;
AffecteGrid(GVte,Action) ;
AffecteGrid(GAch,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GVte) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GAch) ;
GVte.Visible:=False;
GAch.Visible:=False;
end ;

procedure TOM_ARTICLEQTE.OnClose ;
begin
  Inherited ;
  TobVarV.Free; TobVarA.Free;
end ;

procedure TOM_ARTICLEQTE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_ARTICLEQTE.InsertionBaseArtFormule(TobRef : TOB) ;
Var ind : integer;
begin
for ind:=0 to TobRef.Detail.Count-1 do
  begin
  TobRef.Detail[ind].PutValue('GAV_RANG',ind+1);
  TobRef.Detail[ind].PutValue('GAV_ARTICLE',GetField('GAF_ARTICLE'));
  TobRef.Detail[ind].PutValue('GAV_CODEARTICLE',GetField('GAF_CODEARTICLE'));
  end;
if TobRef.Detail.Count > 0 then TobRef.InsertDB(nil);
end ;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TOM_ARTICLEQTE.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp: integer ;
begin
LesCols:=GVte.Titres[0] ;
LesColVte:=LesCols ; icol:=1 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            begin
            if NomCol='GAV_LIBELLEVAR' then
              begin
              VLib:=icol ;
              GVte.ColLengths[icol] := -1;
              end else if NomCol='GAV_VALEURVAR' then
                begin
                VVal:=icol ;
                end else if NomCol='GAV_ENABLEVAR' then
                begin
                VEna:=icol ;
                GVte.ColTypes[icol] := 'B';
                GVte.ColFormats[icol]:=IntToStr(Ord(csCheckbox));
                GVte.FlipBool:=True;
                GVte.ColEditables[iCol]:=false;
                end else if NomCol='GAV_VISIBLEVAR' then
                begin
                VVis:=icol ;
                GVte.ColTypes[icol] := 'B';
                GVte.ColFormats[icol]:=IntToStr(Ord(csCheckbox));
                GVte.FlipBool:=True;
                GVte.ColEditables[iCol]:=false;
                end;
            end;
        end;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

LesCols:=GAch.Titres[0] ;
LesColAch:=LesCols ; icol:=1 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            begin
            if NomCol='GAV_LIBELLEVAR' then
              begin
              ALib:=icol ;
              GAch.ColLengths[icol] := -1;
              end else if NomCol='GAV_VALEURVAR' then
                begin
                AVal:=icol ;
                end else if NomCol='GAV_ENABLEVAR' then
                begin
                AEna:=icol ;
                GAch.ColTypes[icol] := 'B';
                GAch.ColFormats[icol]:=IntToStr(Ord(csCheckbox));
                GAch.FlipBool:=True;
                GAch.ColEditables[iCol]:=false;
                end else if NomCol='GAV_VISIBLEVAR' then
                begin
                AVis:=icol ;
                GAch.ColTypes[icol] := 'B';
                GAch.ColFormats[icol]:=IntToStr(Ord(csCheckbox));
                GAch.FlipBool:=True;
                GAch.ColEditables[iCol]:=false;
                end;
            end;
        end;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

end;

procedure TOM_ARTICLEQTE.ChargeGrille ;
begin
GVte.VidePile(False);
if TobVarV.Detail.Count <= 0 then
  begin
  GVte.Visible:=False;
  end else
  begin
  GVte.Visible:=True;
  TobVarV.PutGridDetail(GVte,False,False,LesColVte,false);
  GVte.RowCount:=TobVarV.Detail.Count+1;
  end;
GAch.VidePile(False);
if TobVarA.Detail.Count <= 0 then
  begin
  GAch.Visible:=False;
  end else
  begin
  GAch.Visible:=True;
  TobVarA.PutGridDetail(GAch,False,False,LesColAch,false);
  GAch.RowCount:=TobVarA.Detail.Count+1;
  end;
end;

Procedure TOM_ARTICLEQTE.RempliTob (Flux : string; RFonct : R_FonctCal);
Var ind : integer;
    TOBRef,TOBL : TOB;
begin
if Flux='VTE' then TOBRef:=TobVarV else TOBRef:=TobVarA;
for ind:=Low(RFonct.VarLibelle) to High(RFonct.VarLibelle) do
  begin
  if RFonct.VarLibelle[ind]<>'' then
    begin
    TOBL:=TOBRef.FindFirst(['GAV_LIBELLEVAR'],[RFonct.VarLibelle[ind]],False);
    if TOBL=Nil then
      begin
      TOBL:=Tob.Create('ARTFORMULEVAR',TOBRef,-1) ;
      TOBL.PutValue('GAV_LIBELLEVAR',RFonct.VarLibelle[ind]);
      TOBL.PutValue('GAV_VENTEACHAT',Flux);
      TOBL.PutValue('GAV_VALEURVAR',0);
      TOBL.PutValue('GAV_ENABLEVAR','X');
      TOBL.PutValue('GAV_VISIBLEVAR','X');
      end ;
    TOBL.AddChampSupValeur('UTILISE','X',False);
    end;
  end;
for ind:=TOBRef.Detail.Count-1 Downto 0 do
  begin
  TOBL:=TOBRef.Detail[ind]  ;
  if (TOBL.FieldExists('UTILISE')) and (TOBL.GetValue('UTILISE')='X') then Continue
                                                                      else TOBL.Free;
  end;

if TOBRef.Detail.Count>0 then TOBRef.Detail[0].DelChampSup('UTILISE',True);
ChargeGrille
end;

Function TOM_ARTICLEQTE.GetTOBLigne (Flux : string; ARow : integer) : TOB ;
Var TOBRef : TOB;
BEGIN
Result:=Nil ;
if Flux='VTE' then TOBRef:=TobVarV else TOBRef:=TobVarA;
if ((ARow<=0) or (ARow>TOBRef.Detail.Count)) then Exit ;
Result:=TOBRef.Detail[ARow-1] ;
END ;

procedure TOM_ARTICLEQTE.TraiterValeur(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
Var TOBL : TOB;
    Flux : string;
    Val : Double;
begin
if GVar.Name='GVTE' then Flux:='VTE' else Flux:='ACH';
TOBL:=GetTOBLigne (Flux,ARow); if TOBL=nil then exit;
ForceUpdate;
Val:=Valeur(GVar.Cells[ACol, ARow]);
TOBL.PutValue('GAV_VALEURVAR',Val);
end;

procedure TOM_ARTICLEQTE.TraiterEnable(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
Var TOBL : TOB;
    Flux,Val : string;
begin
if GVar.Name='GVTE' then Flux:='VTE' else Flux:='ACH';
TOBL:=GetTOBLigne(Flux,ARow); if TOBL=nil then exit;
ForceUpdate;
Val:=UpperCase(Trim(GVar.Cells [ACol, ARow]));
if (Val='X') or (Val='-') then TOBL.PutValue('GAV_ENABLEVAR',Val)
                          else Cancel:=True;
GVar.Cells[ACol, ARow]:=TOBL.GetValue('GAV_ENABLEVAR');
end;

procedure TOM_ARTICLEQTE.TraiterVisible(GVar : THGrid; var ACol, ARow: Integer; var Cancel: Boolean);
Var TOBL : TOB;
    Flux,Val : string;
begin
if GVar.Name='GVTE' then Flux:='VTE' else Flux:='ACH';
TOBL:=GetTOBLigne(Flux,ARow); if TOBL=nil then exit;
ForceUpdate;
Val:=UpperCase(Trim(GVar.Cells [ACol, ARow]));
if (Val='X') or (Val='-') then TOBL.PutValue('GAV_VISIBLEVAR',Val)
                          else Cancel:=True;
GVar.Cells[ACol, ARow]:=TOBL.GetValue('GAV_VISIBLEVAR');
end;

procedure TOM_ARTICLEQTE.SortDuGrid;
Var ACol,ARow : integer ;
    Cancel : boolean ;
begin
ACol:=GVte.Col ; ARow:=GVte.Row ; Cancel:=False ;
GVarCellExit(GVte,ACol,ARow,Cancel) ; if Cancel then Exit ;
GVarRowExit(GVte,ARow,Cancel,False) ; if Cancel then Exit ;
ACol:=GAch.Col ; ARow:=GAch.Row ; Cancel:=False ;
GVarCellExit(GAch,ACol,ARow,Cancel) ; if Cancel then Exit ;
GVarRowExit(GAch,ARow,Cancel,False) ; if Cancel then Exit ;
end;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOM_ARTICLEQTE.GVarExit(Sender: TObject);
begin
SortDuGrid;
end;

procedure TOM_ARTICLEQTE.GVarCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

end;

procedure TOM_ARTICLEQTE.GVarCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var GVar : THGrid;
begin
if Action=taConsult then Exit ;
GVar:=THGrid(Sender) ;
if (ACol=VVal) or (ACol=AVal) then TraiterValeur (GVar,ACol, ARow, Cancel)
  else if (ACol=VEna) or (ACol=AEna) then TraiterEnable (GVar,ACol, ARow, Cancel)
  else if (ACol=VVis) or (ACol=AVis) then TraiterVIsible (GVar,ACol, ARow, Cancel) ;
end;

procedure TOM_ARTICLEQTE.GVarDblClick(Sender: TObject);
Var GVar : THGrid;
    ACol : integer;
begin
if Action=taConsult then Exit ;
GVar:=THGrid(Sender) ;
ACol:=GVar.Col;
if (ACol=VEna) or (ACol=AEna) or (ACol=VVis) or (ACol=AVis) then
  begin
  if GVar.Cells[ACol,GVar.Row]<>'X' then GVar.Cells[ACol,GVar.Row]:='X'
                                    else GVar.Cells[ACol,GVar.Row]:='-';
  end;
end;

procedure TOM_ARTICLEQTE.GVarRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOM_ARTICLEQTE.GVarRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

end;

{==============================================================================================}
{============================= Actions liées aux Boutons ======================================}
{==============================================================================================}
procedure TOM_ARTICLEQTE.AppelFormuleQte(Sender : TObject);
Var Chps,Flux : String;
    RFonct : R_FonctCal;
begin
if TFFiche(Ecran).TypeAction=taConsult then exit ;
if TControl(Sender).name = 'BFORMULEQTEVTE' then Flux:='VTE'
else if TControl(Sender).name = 'BFORMULEQTEACH' then Flux:='ACH'
else exit;
Chps:='GAF_FORMULEQTE'+Flux;
RFonct:=Saisie_FonctionParam(GetField(Chps));
ForceUpdate;
SetField(Chps,RFonct.Formule);
ActivationBoutonFormule;
RempliTob(Flux,RFonct);
end;

procedure TOM_ARTICLEQTE.ActivationBoutonFormule ;
Var BVte,BAch : Boolean;
begin
if TFFiche(Ecran).TypeAction=taConsult then exit ;
BVte:=(GetField('GAF_FORMULEQTEVTE')<>'');
SetControlEnabled('BVTEACH',BVte);
SetControlEnabled('BVTESUP',BVte);
BAch:=(GetField('GAF_FORMULEQTEACH')<>'');
SetControlEnabled('BACHVTE',BAch);
SetControlEnabled('BACHSUP',BAch);
end;

procedure TOM_ARTICLEQTE.CopieFormuleClick(Sender : TObject);
Var Chps1,Chps2,Flux : String;
begin
if TFFiche(Ecran).TypeAction=taConsult then exit ;
if TControl(Sender).name = 'BVTEACH' then
begin
Flux:='VTE';
Chps1:='GAF_FORMULEQTEVTE';
Chps2:='GAF_FORMULEQTEACH';
end else if TControl(Sender).name = 'BACHVTE' then
begin
Flux:='ACH' ;
Chps1:='GAF_FORMULEQTEACH';
Chps2:='GAF_FORMULEQTEVTE';
end else exit;
ForceUpdate;
SetField(Chps2,GetField(Chps1));
ActivationBoutonFormule;
if Flux='VTE' then
  begin
  TobVarA.ClearDetail;
  TobVarA.Dupliquer(TobVarV,True,True);
  TobVarA.PutValueAllFille('GAV_VENTEACHAT','ACH');
  end else
  begin
  TobVarV.ClearDetail;
  TobVarV.Dupliquer(TobVarA,True,True);
  TobVarV.PutValueAllFille('GAV_VENTEACHAT','VTE');
  end;
ChargeGrille;
end;

procedure TOM_ARTICLEQTE.DelFormuleClick(Sender : TObject);
Var Chps,Flux : String;
    RFonct : R_FonctCal;
begin
if TFFiche(Ecran).TypeAction=taConsult then exit ;
if TControl(Sender).name = 'BVTESUP' then Flux:='VTE'
else if TControl(Sender).name = 'BACHSUP' then Flux:='ACH'
else exit;
Chps:='GAF_FORMULEQTE'+Flux;
ForceUpdate;
SetField(Chps,'');
ActivationBoutonFormule;
RFonct:=InitRecord ;
RempliTob(Flux,RFonct);
end;

Initialization
  registerclasses ( [ TOM_ARTICLEQTE ] ) ;
end.
