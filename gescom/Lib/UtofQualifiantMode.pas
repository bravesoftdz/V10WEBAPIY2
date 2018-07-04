unit UtofQualifiantMode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,TarifUtil,DBTables,M3FP,Fe_Main,HDB;

Type
    TOF_QualifiantMode = Class(TOF)
     private
        GridQualifiant : THGRID ;
        TOBQualifiant,TOBQualifiantLigne : Tob ;
        NewLigne: Boolean ;

        procedure GSElipsisClick(Sender: TObject);
        procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure ChargeQualifiant ;
        procedure FormateColSaisie (ACol,ARow : Longint ) ;

        Function VerifQualifiant: boolean ;
        procedure UpdateQualifiant ;
        Function VerifCode(Code: String): boolean ;

        procedure MajDatesTarif ;
        Function  EstRempli( Lig : integer) : Boolean ;


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

     TOF_MAJTarif = Class(TOF)
      public
         procedure OnUpdate ; override ;

     END ;

procedure AGLMajDatesTarif(parms:array of variant; nb: integer ) ;
function CompareTOBEssai (T1,T2: TOB ; ColName : String): Integer ;

const colCode      = 0 ;
      colLibelle   = 1 ;
      colDateDebut = 2 ;
      colDateFin   = 3 ;


      TexteMessage: array[1..6] of string 	= (
        {1}        'Les bornes de dates de validité sont incohérentes'
        {2}        ,'Confirmer vous la suppression de l''enregistrement ?'
        {3}        ,'Confirmez-vous l''abandon des modifications'
        {4}        ,'Le code est obligatoire'
        {5}        ,'Le libellé est obligatoire'
        {6}        ,'Le code que vous avez saisi existe déjà. Vous devez le modifier') ;

implementation

Procedure TOF_QualifiantMode.OnArgument(Arguments:string) ;
var St :String ;
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
GridQualifiant:=THGRID(GetControl('GRIDQUALIFIANT'));
GridQualifiant.OnElipsisClick:=GSElipsisClick  ;
GridQualifiant.OnDblClick:=GSElipsisClick ;
GridQualifiant.OnCellEnter:=GSCellEnter ;
GridQualifiant.OnCellExit:=GSCellExit ;
GridQualifiant.OnRowExit:=GSRowExit ;
GridQualifiant.ColTypes[colDateDebut]:='D' ;
GridQualifiant.ColFormats[colDateDebut]:=ShortdateFormat ;
GridQualifiant.ColTypes[colDateFin]:='D' ;
GridQualifiant.ColFormats[colDateFin]:=ShortdateFormat ;

GridQualifiant.ColFormats[colCode]:='UPPER' ;
GridQualifiant.ColLengths[colCode]:=-1 ;

TFVierge(Ecran).Hmtrad.ResizeGridColumns(GridQualifiant) ;
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
ChargeQualifiant ;
End ;

Procedure TOF_QualifiantMode.OnLoad ;
Begin
inherited ;
TOBQualifiant.PutGridDetail(GridQualifiant,False,False,'',False) ;
End ;

Procedure TOF_QualifiantMode.OnUpdate ;
Begin
inherited ;
SetControlEnabled('BInsert',True) ;
SetControlEnabled('BDelete',True) ;
TOBQualifiant.GetGridDetail(GridQualifiant,TOBQualifiant.Detail.Count+1,'','_CODE;_LIBELLE;_DEBUT;_FIN');
If not VerifQualifiant then exit ;
UpdateQualifiant ;
NewLigne:=False ;
GridQualifiant.ColLengths[colCode]:=-1 ;
End ;

Procedure TOF_QualifiantMode.OnNew ;
Begin
inherited ;
NewLigne:=True ;
GridQualifiant.InsertRow(GridQualifiant.Row) ;
If GridQualifiant.RowCount-1>TOBQualifiant.Detail.count then
  begin
  TOBQualifiantLigne:=Tob.create('un qualifiant',TOBQualifiant,-1);
  TOBQualifiantLigne.AddChampSup('_CODE',False) ;
  TOBQualifiantLigne.AddChampSup('_LIBELLE',False) ;
  TOBQualifiantLigne.AddChampSup('_DEBUT',False) ;
  TOBQualifiantLigne.AddChampSup('_FIN',False) ;
  end ; 
GridQualifiant.Row:=GridQualifiant.Row-1 ;
GridQualifiant.Col:=ColCode ;
GridQualifiant.ColLengths[colCode]:=3 ;
GridQualifiant.SetFocus ;
SetControlEnabled('BInsert',False) ;
SetControlEnabled('BDelete',False) ;
end ;

Procedure TOF_QualifiantMode.OnDelete  ;
begin
inherited ;
if GridQualifiant.Row<=0 then Exit ;
if GridQualifiant.RowCount<=2 then Exit ;
if HShowMessage('1;Qualifiant;Confirmer vous la suppression de l''enregistrement ?;Q;YNC;N;C;','','')<>mrYes then exit;
ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="GCC" AND CC_CODE="'+GridQualifiant.Cells[ColCode,GridQualifiant.Row]+'"') ;
GridQualifiant.DeleteRow(GridQualifiant.row) ;
end;

Procedure TOF_QualifiantMode.OnCancel ;
begin
inherited ;
SetControlEnabled('BInsert',True) ;
SetControlEnabled('BDelete',True) ;
If GridQualifiant.RowCount-1>TOBQualifiant.Detail.count then
   if NewLigne then
     begin
     GridQualifiant.DeleteRow(1) ;
     NewLigne:=False ;
     end ;
OnLoad ;
end ;

Procedure TOF_QualifiantMode.OnClose ;
Var i:Integer ;
TOBQualifiantSav: TOB ;
Compare: integer ;
Begin
inherited ;
Compare:=0 ;
TOBQualifiantSav:= TOB.Create('',NIL,-1);
TOB.Create ('un qualifiant',nil,-1) ;
TOBQualifiantSav.Dupliquer (TOBQualifiant, True, True);
LastError:=0;
if GridQualifiant.rowcount > 1 then TOBQualifiant.GetGridDetail(GridQualifiant,TOBQualifiant.Detail.Count,'un qualifiant','_CODE;_LIBELLE;_DEBUT;_FIN');
For i:=0 to TOBQualifiant.Detail.count-1 do
  begin
  Compare:=CompareTOB (TOBQualifiant.Detail[i],TOBQualifiantSav.Detail[i],'_CODE;_LIBELLE;_DEBUT;_FIN');
  if Compare<>0 then  break ;
  end ;
    if Compare<>0  then
      begin
      If PGIAsk(TexteMessage[3],Ecran.Caption)=mrNo then
          begin
          LastError:=-1;
          Exit;
          end else
          begin
          LastError:=0;
          exit;
          end;
      end;
TOBQualifiant.free ; TOBQualifiant:=nil;
TOBQualifiantSav.free ; TOBQualifiantSav:=nil;
end;



Function TOF_QualifiantMode.VerifQualifiant: Boolean ;
var i: Integer ;
begin
Result:=True ;
If VerifCode(GridQualifiant.Cells[ColCode,GridQualifiant.Row]) then
begin
    LastError:=6; LastErrorMsg:=TexteMessage[LastError] ;
    Result:=False ;
    GridQualifiant.Col := colCode; 
    exit ;
end ;
For i:=0 to TOBQualifiant.Detail.Count-1 do
  begin
  if StrToDate(GridQualifiant.Cells [colDateDebut, i+1]) > StrToDate(GridQualifiant.Cells [colDateFin, i+1]) then
    BEGIN
    LastError:=1; LastErrorMsg:=TexteMessage[LastError] ;
    Result:=False ;
    GridQualifiant.Col := colDateDebut; GridQualifiant.Row := i+1 ;
    break ;
    End ;
  if GridQualifiant.Cells [colCode, i+1] ='' then
    BEGIN
    LastError:=4; LastErrorMsg:=TexteMessage[LastError] ;
    Result:=False ;
    GridQualifiant.Col := colCode; GridQualifiant.Row := i+1 ;
    break ;
    End ;
  if GridQualifiant.Cells [colLibelle, i+1] ='' then
    BEGIN
    LastError:=5; LastErrorMsg:=TexteMessage[LastError] ;
    Result:=False ;
    GridQualifiant.Col := colLibelle; GridQualifiant.Row := i+1 ;
    break ;
    End ;
  end ;
If GridQualifiant.RowCount-1>TOBQualifiant.Detail.count then
  begin
  TOBQualifiantLigne:=Tob.create('un qualifiant',TOBQualifiant,-1);
  TOBQualifiantLigne.AddChampSup('_CODE',False) ;
  TOBQualifiantLigne.AddChampSup('_LIBELLE',False) ;
  TOBQualifiantLigne.AddChampSup('_DEBUT',False) ;
  TOBQualifiantLigne.AddChampSup('_FIN',False) ;
  end ;
End ;

procedure TOF_QualifiantMode.UpdateQualifiant ;
Var CodeQualifiant,Libelle,DateDeb,DateFin,Libre :String ;
i: integer ;
Begin
For i:=0 to TOBQualifiant.Detail.Count-1 do
  Begin
  CodeQualifiant:=TOBQualifiant.Detail[i].GetValue('_CODE') ;
  Libelle:=TOBQualifiant.Detail[i].GetValue('_LIBELLE') ;
  DateDeb:=TOBQualifiant.Detail[i].GetValue('_DEBUT') ;
  DateFin:=TOBQualifiant.Detail[i].GetValue('_FIN') ;
  Libre:=DateDeb+';'+DateFin ;
  If ExisteSQL('Select CC_CODE from CHOIXCOD where CC_TYPE="GCC" and CC_CODE="'+CodeQualifiant+'"') then
  ExecuteSQL('UPDATE CHOIXCOD  SET CC_LIBELLE="'+Libelle+'" , CC_LIBRE="'+Libre+'" where CC_CODE="'+CodeQualifiant+'"')
  Else ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_LIBRE) VALUES ("GCC","'+CodeQualifiant+'","'+Libelle+'","'+Libre+'")') ;
  end ;
end ;

Function TOF_QualifiantMode.EstRempli( Lig : integer) : Boolean ;
var i : integer ;
BEGIN
Result:=False ;
for i:=0 to GridQualifiant.ColCount-1 do
  if (GridQualifiant.Cells[i,Lig]<>'') then begin result:= true; break; end;
END ;

{Procedure TOF_QualifiantMode.NouvelleLigne ;
Begin
end ; }



Function TOF_QualifiantMode.VerifCode(Code: String): boolean ;
Var Q: TQuery ;
begin
Q:=OpenSql('Select CC_CODE From ChoixCod Where CC_TYPE="GCC" And CC_CODE="'+Code+'"',True) ;
Result:=Not Q.Eof ; Ferme(Q) ;
end ;

{==============================================================================================}
{=============================== Evenement du Grid ========================================}
{==============================================================================================}
procedure TOF_QualifiantMode.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_F5 : if (Screen.ActiveControl=GridQualifiant) then GSElipsisClick(Sender);
    END;
end;

procedure TOF_QualifiantMode.GSElipsisClick(Sender: TObject);
Var DATE: THCritMaskEdit;
    Coord : TRect;
begin
     Inherited ;
if (GridQualifiant.Col = colDateDebut) or (GridQualifiant.Col = colDateFin) then
    BEGIN
    Coord := GridQualifiant.CellRect (GridQualifiant.Col, GridQualifiant.Row);
    DATE := THCritMaskEdit.Create (GridQualifiant);
    DATE.Parent := GridQualifiant;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then GridQualifiant.Cells[GridQualifiant.Col,GridQualifiant.Row]:= DATE.Text;
    DATE.Destroy;
    END;
END;

procedure TOF_QualifiantMode.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if Not EstRempli(ou) then
begin
     GridQualifiant.DeleteRow(ou) ;
     SetControlEnabled('BInsert',True) ;
     SetControlEnabled('BDelete',True) ;
end ;
end;

procedure TOF_QualifiantMode.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
If Action=taConsult then Exit ;
GridQualifiant.ElipsisButton := ((GridQualifiant.Col = colDateDebut) or (GridQualifiant.Col = colDateFin)) ;
End ;

procedure TOF_QualifiantMode.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
If Action=taConsult then Exit ;
FormateColSaisie(Acol, Arow) ;
If (Acol=ColCode) then
  if(GridQualifiant.Cells[ColCode,Arow]='') then GridQualifiant.Col:=ColCode ;
If (Acol=ColLibelle) and (GridQualifiant.Cells[ColLibelle,Arow]='')  then GridQualifiant.Col:=ColLibelle ;
End ;

{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}
procedure TOF_QualifiantMode.FormateColSaisie (ACol,ARow : Longint ) ;
Var st, Stc:String ;
Begin
St:=GridQualifiant.Cells[ACol,ARow] ; StC:=St ;
if ACol=0 then StC:=Trim(St);
GridQualifiant.Cells[ACol,ARow]:=StC ;
end ;



{==============================================================================================}
{=============================== Chargement des qualifiant ====================================}
{==============================================================================================}
procedure TOF_QualifiantMode.ChargeQualifiant ;
var QQualif: TQuery ;
Date,DateDebut,DateFin: String ;
Begin
QQualif:=OpenSql('Select * from ChoixCod where CC_TYPE="GCC"',True) ;
if not QQualif.EOF then
begin
  TOBQualifiant:= TOB.Create('les qualifiants',NIL,-1);
  While not QQualif.EOF do
     Begin
     TOBQualifiantLigne:=Tob.create('un qualifiant',TOBQualifiant,-1);
     TOBQualifiantLigne.AddChampSup('_CODE',False) ;
     TOBQualifiantLigne.AddChampSup('_LIBELLE',False) ;
     TOBQualifiantLigne.AddChampSup('_DEBUT',False) ;
     TOBQualifiantLigne.AddChampSup('_FIN',False) ;
     TOBQualifiantLigne.PutValue('_CODE',QQualif.Findfield('CC_CODE').AsString);
     TOBQualifiantLigne.PutValue('_LIBELLE',QQualif.Findfield('CC_LIBELLE').AsString);
     // Recup des dates
     Date:=QQualif.FindField('CC_LIBRE').AsString ;
     if Date='' then  exit ;
     DateDebut:=ReadTokenSt(Date) ;
     DateFin:=Date ;
     TOBQualifiantLigne.PutValue('_DEBUT',DateDebut);
     TOBQualifiantLigne.PutValue('_FIN',DateFin);

     QQualif.next ;
     End ;
     ferme(QQualif) ;
end ;
GridQualifiant.RowCount:=TOBQualifiant.Detail.count+1 ;
end;

{==============================================================================================}
{=============================== Mise à jour tarif ============================================}
{==============================================================================================}
procedure TOF_QualifiantMode.MajDatesTarif ;
Begin
AGLLanceFiche('GC','GCMAJTARIFMODE','','','') ;
end ;

procedure AGLMajDatesTarif( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTof else exit;
if (TOTOF is TOF_QualifiantMode) then TOF_QualifiantMode(TOTOF).MajDatesTarif else exit;
end;

{==============================================================================================}
{=============================== TOF MAJTarif =================================================}
{==============================================================================================}
procedure TOF_MAJTarif.OnUpdate ;
Var Qualifiant: String ;
Debut,Fin: TDate ;
Begin
Qualifiant:=GetControlText('_QUALIFIANT') ;
Debut:=StrToDate(GetControlText('_DEBUT')) ;
Fin:=StrToDate(GetControlText('_FIN')) ;
ExecuteSQL('UPDATE TARIF  SET GF_DATEDEBUT="'+USDATETIME(Debut)+'" , GF_DATEFIN="'+USDATETIME(Fin)+'" where GF_TARFMODE="'+Qualifiant+'"') ;
End ;

function CompareTOBEssai (T1,T2: TOB ; ColName : String): Integer ;
Var v1,v2 : Variant ;
    j,ier : integer ;
    ch : string ;
BEGIN
result:=0 ;
if T1.NomTable<>T2.NomTable then exit ;
While (ColName<>'') and (Result=0) do
  BEGIN
  Ch:=ReadTokenSt(ColName) ;
  if (Ch<>'') and (Ch[1]='-') then BEGIN Delete(Ch,1,1) ; ier:=-1 ; END else ier:=1 ;
j:=T1.GetNumChamp(Ch) ;
  if j>0 then
     BEGIN
     V1:=T1.GetValeur(j) ; V2:=T2.GetValeur(j) ;
     Case VarType(V1) of
       varNull    : ;
       varDate    : BEGIN if V1<V2 then Result:=ier*-1 else if V1>V2 then Result:=ier*1 ; END ;
       varByte,varDouble,varSingle,varSmallint,varInteger,varCurrency  : BEGIN if V1<V2 then Result:=ier*-1 else if V1>V2 then Result:=ier*1 ; END ;
       else         result:=ier*CompareText(v1,v2) ;
       END ;
     END ;
  END ;
END ;



Initialization
registerclasses([TOF_QualifiantMode]);
registerclasses([TOF_MAJTarif]);
RegisterAglProc( 'MajDatesTarif', TRUE , 0, AGLMajDatesTarif );
end.
