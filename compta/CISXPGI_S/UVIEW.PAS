{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /
Description .. : Unité permettant de visualiser les données générées à partir 
Suite ........ : des fichiers en appliquant le script.
Mots clefs ... : 
*****************************************************************}

unit Uview;

interface

uses
	Windows, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
	Grids, DBGrids, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB,{$ENDIF}
  ExtCtrls,
	ASCIIV, StdCtrls, HTB97, Menus, Buttons, HCTRLS,
{$IFNDEF EAGLCLIENT}
        PrintDbG,
{$ENDIF}
        HXlspas,
        Hent1, HPanel;

type
	TViewForm = class(TForm)
                DataSource1: TDataSource;
                Panel2: TPanel;
                DockBottom: TDock97;
                WhatError: TLabel;
                Uutils97: TToolbar97;
                BImprimer: TToolbarButton97;
                BRecherche: TToolbarButton97;
                ToolbarSep971: TToolbarSep97;
                BMenuZoom: TToolbarButton97;
                Valide97: TToolbar97;
                BValider: TToolbarButton97;
                BFerme: TToolbarButton97;
                BAide: TToolbarButton97;
                POPZ: TPopupMenu;
                FindDialog: TFindDialog;
                HPanel1: THPanel;
                DBGrid1: TDBGrid;
                PCumul: TPanel;
                ASCIIView1: TASCIIView;
                Exporte: TToolbarButton97;
                procedure FormCreate(Sender: TObject);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                procedure FormDeactivate(Sender: TObject);
                procedure FormDestroy(Sender: TObject);
                procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
                procedure BMenuZoomMouseEnter(Sender: TObject);
                procedure BImprimerClick(Sender: TObject);
                procedure BRechercheClick(Sender: TObject);
                procedure FindDialogFind(Sender: TObject);
                procedure DBGrid1ColumnMoved(Sender: TObject; FromIndex,
                  ToIndex: Integer);
                procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
                  DataCol: Integer; Column: TColumn; State: TGridDrawState);
                procedure BAideClick(Sender: TObject);
                procedure ExporteClick(Sender: TObject);
                private
		            { Déclarations privées }
		            FTableName         : String;
		            FDatabaseName      : String;
                ListeSql           : HTStringList;
                ListeOrder         : HTStringList;
                OkDecoup           : Boolean;
		            Activated          : Boolean;
		            ClosingItSelf      : Boolean;
                Q                  : TQuery;
                FirstFind          : boolean;
                ModeAsc            : integer;
                StCumul,OrdreDeTri : string;
                procedure BMenuZoomClick(Sender: TObject);
                public
                        TravailEnCours : Boolean;
                        { Déclarations publiques }
                constructor Create(ATableName,ADatabaseName : string; ListeEnreg : HTStringList; Decoupage : Boolean; ModeASCII : integer; OrderBy : string=''; Lorder : HTStringList=nil); reintroduce; overload ; virtual ;
                procedure CreateParams(var Params: TCreateParams); override;
                procedure PopZoom97 ( POPZ : TPopupMenu ) ;
                function  CreerLigPop ( Name : string; Owner : TComponent) : TMenuItem ;
                procedure MajZoneCumul(FCumulField : string);
                procedure UpdateZoneCumul;
	end;
var
	ViewForm: TViewForm;

procedure ScriptVisu (Name,	DatabaseName : string; ListeEnreg : HTStringList; Decoupage : Boolean=FALSE; ModeASCII : integer=0; OrderBy :string=''; ListeOrder : HTStringList=nil);

implementation

{$R *.DFM}

uses UDMIMP;

procedure ScriptVisu (Name, DatabaseName : string; ListeEnreg : HTStringList; Decoupage : Boolean=FALSE; ModeASCII : integer=0; OrderBy :string=''; ListeOrder : HTStringList=nil);
var XX : TViewForm ;
BEGIN
    XX:= TViewForm.Create(Name, DataBaseName, ListeEnreg, Decoupage, ModeASCII, OrderBy, ListeOrder) ;
    XX.ShowModal ;
    XX.free;
end;

type
	TAssociation = record
		F : TStringField;
		T : TTable;
	end;


(* ---------------------------------------------------------------------- *)
(*                                                                        *)
(* ---------------------------------------------------------------------- *)
constructor TViewForm.Create(ATableName, ADatabaseName  : string; ListeEnreg : HTStringList; Decoupage : Boolean; ModeASCII : integer; OrderBy :string=''; Lorder : HTStringList=nil);
begin
	FTableName    := ATableName;
	FDatabaseName := ADatabaseName;
        ListeSQl      := ListeEnreg;
        OkDecoup      := Decoupage;
        ModeAsc       := ModeAscii;
        OrdreDeTri    := OrderBy;
        ListeOrder    := Lorder;
	inherited Create(Application);
end;

procedure TViewForm.CreateParams(var Params: TCreateParams);
begin
	Caption := FTableName;
	inherited CreateParams(Params);
end;

procedure TViewForm.UpdateZoneCumul;
var
j, i, iLeft, iFind : integer;
HN                 : THNumEdit ;
begin
      if DBGrid1.Columns.Count > 15 then exit;
      iFind := 0;
      for j:=0 to PCumul.ControlCount-1 do
       if PCumul.Controls[j] is THNumEdit then
       BEGIN
         HN:=THNumEdit(PCumul.Controls[j]) ;
         iLeft:=DBGrid1.Left+Ord(dgIndicator in DBGrid1.Options)*(IndicatorWidth+1) ;

         for i:=0 to DBGrid1.Columns.Count-1 do
                if Uppercase(DBGrid1.Columns[i].FieldName)=Uppercase(HN.Name) then
                BEGIN
                   iFind:=i ;
                   break ;
                END
                else iLeft:=iLeft + DBGrid1.Columns[i].Width+1;
                if iFind>=0 then
                begin
                    HN.Left:=iLeft;
                    HN.Width:=DBGrid1.Columns[ifind].Width +1;
                end;
       END ;
end;

procedure TViewForm.MajZoneCumul(FCumulField : string);
Var Q1 : TQuery ;
    stCount,StTable,StSQL,Ch,St,masks : String ;
    HN : THNumEdit ;
    jj,nbc,j,Max : integer ;
begin
if DBGrid1.Columns.Count > 15 then exit;
PCumul.Visible:=FCumulField<>'' ;
if Not PCumul.Visible then exit ;
St:=FCumulField ;
StTable:=FTableName;//+'.DB' ;
StTable:=FindEtReplace(StTable,',',';',TRUE)+';' ;
StTable:=TRIM(Uppercase(ReadTokenSt(StTable))) ;
Masks:='#,##0' ; jj:=-1 ; nbc:=0;
While PCumul.ControlCount>0 do PCumul.Controls[0].Free ;
stCount:='';
StSQL := 'SELECT ';
While St<>'' do
   BEGIN
       Ch:=TRIM(Uppercase(ReadTokenSt(St))) ;
       inc(nbc);
       if Ch<>'' then
       begin
            if St <> '' then
               StSQL:=StSQL+'SUM('+Ch+'), '
            else
               StSQL:=StSQL+'SUM('+Ch+') ';
            Inc(jj) ;
       end;
       HN:=THNumEdit.Create(PCumul.Owner) ;
       HN.Parent:=PCumul ;
       HN.ParentColor:=TRUE  ; HN.Font.Style:=[fsBold] ;
       HN.Value:=0 ;
       HN.Tag := nbc;
       HN.Top:=-1 ; HN.Ctl3D:=FALSE ;
       HN.Name := Ch ;
   END ;
    Max:=jj+1 ;
    PCumul.Tag:=Max ;
    try
      if  not DMImport.DBGlobalD.Connected then
      DMImport.DBGlobalD.Connected := TRUE ;
      Q1 := OpenSQLADO (StSQL+ 'FROM ' + StTable + OrdreDeTri, DMImport.DBGlobalD.ConnectionString);
      DataSource1.DataSet := Q;
    except
     Q1:=Nil ;
    end ;
   if Q1=Nil then PCumul.Visible:=FALSE else
   if (Not Q1.EOF) then
      BEGIN
      for j:=0 to PCumul.ControlCount-1 do
       if PCumul.Controls[j] is THNumEdit then
         BEGIN
         HN:=THNumEdit(PCumul.Controls[j]) ;
         if HN.Tag>0 then
            HN.Value:=Q1.Fields[HN.Tag-1].AsFloat ;
         END ;
       PCumul.Caption:=TraduireMemoire('Totaux') ;
      END ;
   if (Q1<>Nil) then Ferme (Q1);
   UpdateZoneCumul;
end;

procedure TViewForm.FormCreate(Sender: TObject);
var N               : Integer;
    iet,iPoint      : Integer;
    FieldName,Chaine: string;
    Oktable         : Boolean;
    SQL             : string;
begin
    if ModeAsc = 1 then
    begin
       Panel2.Visible         := false;
       DBGrid1.Visible        := false;
       ASCIIView1.Align       := Alclient;
       ASCIIView1.FileName    := FTableName;
       exit;
    end
    else
    begin
         ASCIIView1.visible   := FALSE;
         DBGrid1.Visible      := TRUE;
    end;
    if not OkDecoup then
    SQL := 'SELECT * from '+ FTableName + OrdreDeTri
    else
    begin
          if ListeOrder<>nil then
          begin
           for N:= 0 to ListeOrder.count-1 do
               if pos(ListeSql[ListeSql.count-1], ListeOrder[N]) <> 0 then
               begin
                    OrdreDeTri := ' Order By ' +  ListeOrder[N];
                    break;
               end;
          end;
         SQL := 'SELECT * from '+ FTableName+ListeSql[ListeSql.count-1] + OrdreDeTri;
    end;
    if  not DMImport.DBGlobalD.Connected then
     DMImport.DBGlobalD.Connected := TRUE ;
    Q := OpenSQLADO (SQL, DMImport.DBGlobalD.ConnectionString);
    DataSource1.DataSet := Q;
    try
    if not OkDecoup then
    begin
        ListeSql := HTStringList.Create;

		with DBGrid1 do
			for N:=0 to DBGrid1.FieldCount-1 do
                        begin
                         // remplissage de la liste SQL
                         FieldName := Fields[N].FieldName;
                         if Fields[N].DataType = ftFloat then
                         begin
                              if StCumul = '' then
                              StCumul := FieldName
                              else
                              StCumul := StCumul + ';' +FieldName;
                         end;
                         iPoint := pos('_', FieldName);
                         if (iPoint <> 0) then
                         begin
                              FieldName := copy(Fields[N].FieldName, 1, iPoint-1);
                              Oktable := FALSE;
                              for iet := 0 to ListeSql.count-1 do
                              begin
                                   Chaine := copy (ListeSql[iet], 1, iPoint-1);
                                   if Chaine = FieldName then begin Oktable := TRUE; break; end;
                              end;
                              if not Oktable then
                              ListeSql.add (Fields[N].FieldName)
                              else
                              ListeSql.Strings[iet] := ListeSql.Strings[iet] +','+Fields[N].FieldName;
                         end;
                         // fin
                         if Fields[N] is TCurrencyField then
                            Columns[N].Width := Columns[N].Width + 40;
                        end;
    end;
    BMenuZoomMouseEnter(Sender);
    except
		on e: EDatabaseError do
			Application.ShowException(e);
    end;
                              MajZoneCumul (StCumul);
end;

(* ------------------------------------------------------------------------ *)
(*                                                                          *)
(* ------------------------------------------------------------------------ *)

procedure TViewForm.FormDestroy(Sender: TObject);
var AHandle : THandle;
begin	(* Release *)
	if ClosingItSelf then begin
		AHandle := TWinControl(Owner).Handle;
	end;
end;

procedure TViewForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Ferme (Q);
	ClosingItSelf := True;
	Action := caFree;	(* pour liberer la grille *)
    if (not OkDecoup) and (ListeSQl <> nil) then ListeSql.destroy;
end;

(* ------------------------------------------------------------------------ *)
(*                                                                          *)
(* ------------------------------------------------------------------------ *)

procedure TViewForm.FormDeactivate(Sender: TObject);
begin
	Activated := False;
end;

procedure TViewForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
	CanClose := not TravailEnCours;
end;

procedure PurgePopup( PP : TPopupMenu ) ;
Var M,N : TMenuItem ;
BEGIN
if PP=Nil then Exit ;
if PP.Items.Count<=0 then Exit ;
While PP.Items.Count>0 do
   BEGIN
   M:=PP.Items[0] ;
   While M.Count>0 do BEGIN N:=M.Items[0] ; M.Remove(N) ; N.Free ; END ;
   PP.Items.Remove(M) ; M.Free ;
   END ;
END ;

Function IsButtonPop ( CC : TComponent ) : boolean ;
BEGIN
Result:=True ;
if CC.ClassType=TBitBtn then Exit ;
if CC.ClassType=TSpeedButton then Exit ;
if CC.ClassType=TToolbarButton97 then Exit ;
Result:=False ;
END ;

function TViewForm.CreerLigPop ( Name : string; Owner : TComponent) : TMenuItem ;
Var
    T   : TMenuItem ;
BEGIN
T         := TMenuItem.Create(Owner) ;
T.Name    := Name ;
T.Caption := T.Name;
T.OnClick := BMenuZoomClick;
Result    := T ;
END ;

procedure TViewForm.PopZoom97 ( POPZ : TPopupMenu ) ;
Var
    T     : TMenuItem ;
    iet   : integer;
    ip    : integer;
BEGIN
     PurgePopup(POPZ) ;
     For iet:=0 To ListeSql.Count-1 Do
     begin
                 ip := pos('_', ListeSql[iet]);
                 if (ip <> 0) then
                      T:=CreerLigPop(copy(ListeSql[iet], 1, ip-1),POPZ)
                 else
                      T:=CreerLigPop(ListeSql[iet],POPZ) ;
                 POPZ.Items.Add(T) ;
     end;
END ;

procedure TViewForm.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(POPZ) ;
end;

procedure TViewForm.BMenuZoomClick(Sender: TObject);
var
  MI      : TMenuItem;
  Name    : string;
  iet     : integer;
  SQl     : string;
  SelectQL: string;
begin
  MI   := Sender as TMenuItem;
  Name := MI.Name;
  SQl  := '';
  if not OkDecoup then
  begin
       For iet:=0 To ListeSql.Count-1 Do
       begin
                 if Copy(ListeSql[iet], 1, length(Name)) = Name then
                 begin
                      SQl :=  ListeSql[iet];
                      break;
                 end;
       end;
       Name := '';
  end else
  begin
       SQL := '*'; OrdreDeTri := '';
       if ListeOrder<>nil then
       begin
            for iet:= 0 to ListeOrder.count-1 do
            begin
               if Copy(ListeOrder[iet], 1, length(Name)) = Name then
               begin
                    OrdreDeTri := ' Order By ' + ListeOrder[iet]; break;
               end;
            end;
       end;
  end;
 if SQl <> '' then
  begin
  SelectQL := ' SELECT '+SQL+' from '+FTableName+Name + OrdreDeTri;
  Ferme (Q);
  if  not DMImport.DBGlobalD.Connected then
     DMImport.DBGlobalD.Connected := TRUE ;
  Q := OpenSQLADO (SelectQL, DMImport.DBGlobalD.ConnectionString);
  DataSource1.DataSet := Q;
  end;
end;

procedure TViewForm.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
     PrintDBGrid(DbGrid1, nil, Caption, '');
{$ENDIF}
end;

procedure TViewForm.BRechercheClick(Sender: TObject);
begin
     FirstFind := true;
     FindDialog.Execute;
end;

procedure TViewForm.FindDialogFind(Sender: TObject);
begin
     Rechercher(DbGrid1, FindDialog, FirstFind);
end;


procedure TViewForm.DBGrid1ColumnMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
begin
     MajZoneCumul (StCumul);
end;

procedure TViewForm.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
     UpdateZoneCumul;
end;

procedure TViewForm.BAideClick(Sender: TObject);
begin
     CallHelpTopic(Self) ;
end;

procedure TViewForm.ExporteClick(Sender: TObject);
var SD : TSaveDialog;
begin
{$IFNDEF EAGLCLIENT}
(*  SD := TSaveDialog.Create(Application);
  SD.DefaultExt := '*.xls';
  SD.Title := 'Exportation de fichiers';
  SD.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.csv)|*.csv';
  if SD.execute then
     ExportGrid(nil, DBGrid1, SD.FileName, SD.FilterIndex, TRUE) ;
  SD.free;
*)
{$ENDIF}
end;

end.
