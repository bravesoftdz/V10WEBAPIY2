{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : QBPCALENDREP (QBPCALENDREP)
Mots clefs ... : TOM;QBPCALENDREP
*****************************************************************}
Unit QBPCALENDREP_TOM ;

Interface

Uses {$IFNDEF EAGLCLIENT} db, DBCtrls,
     {$ELSE} UtileAGL,UTob, {$ENDIF}
     UTOM, Ugraph,HEnt1,HCtrls;

Type
  TOM_QBPCALENDREP = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
     OkNew,OkValid:boolean;
     code:string;
     QT: TQTob;
     ValCell:double;
     days: array[1..7] of string;
     LigneSolde,ColSolde:integer;
     OkModifCellule:boolean;
     procedure BValider_OnClick(Sender: TObject);
     procedure BtnSolde(Sender: TObject);
     procedure RecalculTot(F:THGrid);
     procedure MAJTableGrid;
     procedure RemplitGrid;
     procedure CellExit(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);
     procedure CellEnter(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);
    end ;

Implementation

Uses Classes,Sysutils,HTB97,
     {$IFNDEF EAGLCLIENT}Fiche,{$ELSE}efiche,{$ENDIF}
    UUtil;

procedure TOM_QBPCALENDREP.BValider_OnClick(Sender: TObject);
begin
  inherited;
  { EVI / Force le statut de la fiche pour passer dans le UpdateRecord }
  if TFFiche( Ecran ).FTypeAction <> taCreat then TFFiche( Ecran ).QFiche.Edit;
  TFFiche(Ecran).Bouge(nbPost) ;
  Ecran.Close ;
end;


procedure TOM_QBPCALENDREP.RecalculTot(F:THGrid);
var i:Integer ;
    somme:double ;
begin
 somme:=0 ;
 for i:=0 to F.RowCount-1 do
  somme:=somme+Valeur(F.Cells[1,i]) ;
 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));
end ;

procedure TOM_QBPCALENDREP.OnNewRecord ;
begin
  Inherited ;
  OkNew:=true;
  RemplitGrid;
end ;

procedure TOM_QBPCALENDREP.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_QBPCALENDREP.OnUpdateRecord ;
begin
  Inherited ;
  MAJTableGrid;
  if OkValid then OkModifCellule:=false;
end ;

procedure TOM_QBPCALENDREP.OnAfterUpdateRecord ;
begin
  Inherited ;
  if OkValid
   then MExecuteSql('UPDATE QBPCALENDREP SET '+code+
                   ' WHERE QBN_CALENDREP="'+getfield('QBN_CALENDREP')+'"',
                   'TOM_QBPCALENDREP (OnAfterUpdateRecord).');
end ;

procedure TOM_QBPCALENDREP.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_QBPCALENDREP.OnLoadRecord ;
begin
  Inherited ;
  OkValid:=true;
  RemplitGrid;
end ;

procedure TOM_QBPCALENDREP.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_QBPCALENDREP.OnArgument ( S: String ) ;
begin
  Inherited ;
  TToolbarButton97(GetControl('BValider')).OnClick := BValider_OnClick;
  THGrid(getcontrol('GRIDLOI')).ColFormats[1]:='#,####0.0000';
  THGrid(getcontrol('GRIDLOI')).ColTypes[1]:='F';
  OkNew:=false;
  TToolBarButton97(getcontrol('BTNSOLDE')).OnClick := BtnSolde;
  THGrid(getcontrol('GRIDLOI')).OnCellExit:=CellExit;
  THGrid(getcontrol('GRIDLOI')).OnCellEnter:=CellEnter;

  days[1] := TraduireMemoire('Lundi');
  days[2] := TraduireMemoire('Mardi');
  days[3] := TraduireMemoire('Mercredi');
  days[4] := TraduireMemoire('Jeudi');
  days[5] := TraduireMemoire('Vendredi');
  days[6] := TraduireMemoire('Samedi');
  days[7] := TraduireMemoire('Dimanche');

  OkModifCellule:=false;
end ;

procedure TOM_QBPCALENDREP.OnClose ;
begin
 Inherited ;
 if not OkValid
  then LastError:=-1
  else
   begin
    LastError:=0;
    //suppression de la tob
    if QT <> nil
     then QT.Free;
   end;
  LastErrorMsg:=TraduireMemoire('La somme des pourcentages doit être égale à 100.');
end ;

procedure TOM_QBPCALENDREP.OnCancelRecord ;
begin
  Inherited ;  
end ;



procedure TOM_QBPCALENDREP.RemplitGrid;
var i:integer;
    somme:double;
begin
 if QT <> nil
  then QT.Free;
 QT :=TQTob.create(['LADATE'],['VAL1']);
 THGrid(getcontrol('GRIDLOI')).RowCount:=0;

 somme:=0;

 if OkNew
  then
   begin
    for i:=1 to 7 do
     setField('QBN_PRCT'+IntToStr(i),0);
   end;

   for i:=1 to 7 do
    begin
     QT.addValeur([Days[i]], [getfield('qbn_prct'+IntToStr(i))]);
     somme:=somme+getfield('qbn_prct'+IntToStr(i));
    end;
   QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');

 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));
end;



procedure TOM_QBPCALENDREP.MAJTableGrid;
var i:integer;
    {$IFDEF MODE} val:hString; //EVI_TEMP_V800
    {$ELSE} val:String; {$ENDIF}
    somme:double;
begin
 code:='';
 somme:=0;

 for i := 0 to QT.laTob.Detail.Count-1 do
  begin
   QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'LADATE;VAL1');
   val:=QT.laTob.getValue('VAL1');
   somme:=somme+StrToFloat(val);
   RemplaceVirguleParPoint(val);
   if code=''
    then code:=' QBN_PRCT'+IntToStr(i+1)+'="'+val+'" '
    else code:=code+', QBN_PRCT'+IntToStr(i+1)+'="'+val+'" ';
  end;
 OkValid:=true;
 if ((VALEUR(Format('%10.4f',[somme]))=100) or (VALEUR(Format('%10.4f',[somme]))=0))=false
  then OkValid:=false;
end;

procedure TOM_QBPCALENDREP.BtnSolde(Sender: TObject);
var val,somme:double;
    i:integer;
begin
 MAJTableGrid;
 QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneSolde,'LADATE;VAL1');
 val:=QT.laTob.getValue('VAL1');
 QT.laTob.SetDouble('VAL1',val+StrToFloat(GetControlText('EDTSOLDE')));
 QT.laTob.PutLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneSolde,false,false,'LADATE;VAL1');
 somme:=0;  
 for i := 0 to QT.laTob.Detail.Count-1 do
  begin
   QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'LADATE;VAL1');
   val:=QT.laTob.getValue('VAL1');
   somme:=somme+val;
  end;
 
 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));

end;

procedure TOM_QBPCALENDREP.CellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 LigneSolde:=G.row;
 ColSolde:=G.col;
 ValCell:=StrToFloat(G.cells[ColSolde,LigneSolde]);
end;

procedure TOM_QBPCALENDREP.CellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 RecalculTot(G) ;
 OkModifCellule:=true;
end;

Initialization
  registerclasses ( [ TOM_QBPCALENDREP ] ) ; 
end.

