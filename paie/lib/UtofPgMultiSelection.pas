unit UtofPgMultiSelection;

interface

uses Classes,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}SysUtils,
     Utof,Utob,HCtrls,Htb97,HmsgBox,Vierge;

type
  TOF_PGMULTISELECTION = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate;                      override;
    procedure OnClose;                       override;
    procedure OnDelete;                      Override;
  private
   Tob_Select,Tob_Result,Tob_Delete : Tob;
   Tablette,Champ, GblTypeAction : string;
   GDSelect,GDResult : THGrid;
   Procedure ClickAPrendre(Sender : TObject);
   Procedure ClickARendre(Sender : TObject);
   Procedure ChangeGridObjet(Sender , Source : TObject ; Tob_Sender, Tob_Source : Tob);
   Procedure AfficheGrid;
  end;

implementation

{ TOF_PGMULTISELECTION }



procedure TOF_PGMULTISELECTION.OnArgument(Arguments: string);
Var Q : TQuery;
    TMere,TFille,T : Tob;
    i : Integer;
    Btn : TToolBarButton97;
    St,St2,Valeur,Table,Code,Libelle : String;
begin
  inherited;
  Valeur := ReadTokenSt(Arguments);
  Champ := ReadTokenSt(Arguments);
  Tablette := ReadTokenSt(Arguments);
  GblTypeAction := '';
  GDSelect := THGrid(GetControl('GDSELECT'));
  GDResult := THGrid(GetControl('GDRESULT'));

  Tob_Select:=Tob.Create('Les sélectons',nil,-1);
  T := Tob.Create('Les sélectons',Tob_Select,-1);
  T.AddChampSupValeur('CODE','TOUS');
  T.AddChampSupValeur('LIBELLE','Tous');
  Tob_Result:=Tob.Create('Les résultants',nil,-1);

  St:='';
  Q := OpenSql('SELECT * FROM DECOMBOS WHERE DO_COMBO="'+Tablette+'"',True);
  if Not Q.Eof then
    Begin
    Code :=Q.FindField('DO_CODE').AsString;
    Libelle := Q.FindField('DO_CHAMPLIB').AsString;
    Table := PrefixeToTable(Q.FindField('DO_PREFIXE').ASString);
    St := 'SELECT '+Code+','+Libelle+
          ' FROM '+Table;
    if Q.FindField('DO_WHERE').ASString<>'' then
      Begin
      St2 := Q.FindField('DO_WHERE').ASString;
      If Pos('&#@',St2)>0 then
        If Pos('&#@',St2)=1 then
          St2:=Trim(Copy(St2,Pos('&#@',St2)+3,length(St2)))
        Else
          St2:=Trim(Copy(St2,1,Pos('&#@',St2))+Copy(St2,Pos('&#@',St2)+3,length(St2)));
      if Trim(St2)<>'' then St := St + ' WHERE '+St2;
      End;
    End;
  Ferme(Q);
  IF St<>'' then
    Begin
    Q := OpenSQL(St,True);
    If Not Q.eof then
       Begin
       TMere := Tob.Create('Temp tablette',nil,-1);
       TMere.LoadDetailDB('Temp tablette','','',Q,False);
       End;
    Ferme(Q);
    End;
  FreeAndNil(TFille);
  i := 0;
  if Assigned(TMere) then TFille := TMere.detail[0];
  While Assigned(TFille) Do
    Begin
//    If not ((TFille.GetValue('DO_PREFIXE')='CO') OR (TFille.GetValue('DO_PREFIXE')='CC')) then Break;
    T := Tob.Create('Les sélectons',Tob_Select,-1);
    T.AddChampSupValeur('CODE',TFille.GetValue(Code));
    T.AddChampSupValeur('LIBELLE',TFille.GetValue(Libelle));
    Inc(i);
    if i >= TMere.detail.count then Break;
    TFille:= TMere.detail[i];
    End;
  FreeAndNil(TMere);


  Q:=OpenSql('SELECT * FROM PGMULTISELECTION WHERE PMZ_MULTISELECTION="'+Valeur+'"', True);
  If Not Q.Eof then
    Begin
     Tob_Delete := Tob.Create('PGMULTISELECTION',nil,-1);
     Tob_Delete.LoadDetailDB('PGMULTISELECTION','','',Q,False);
    End
  else
     GblTypeAction := 'CREATION';
  Ferme(Q);

   If (Assigned(Tob_Select)) and (Assigned(Tob_Delete)) then
     if (Tob_Select.detail.count>0) AND (Tob_Delete.detail.count>0) Then
        Begin
        SetControlText('PMZ_MULTISELECTION',Tob_Delete.detail[0].GetValue('PMZ_MULTISELECTION'));
        SetControlEnabled('PMZ_MULTISELECTION',False);
        SetControlText('PMZ_LIBELLE',Tob_Delete.detail[0].GetValue('PMZ_LIBELLE'));
        For i:=0 to Tob_Delete.detail.count-1 do
           Begin
           T :=Tob_Delete.detail[i];
           TFille := Tob_Select.FindFirst(['CODE'],[T.GetValue('PMZ_VALEUR')],False);
           if Assigned(TFille) then
             Begin
             TFille.ChangeParent(Tob_Result,-1);
             End;
           End;
       End;

  AfficheGrid;


Btn := TToolBarButton97(GetControl('APRENDRE'));
if Assigned(Btn) then Btn.OnClick:=ClickAPrendre;
Btn := TToolBarButton97(GetControl('ARENDRE'));
if Assigned(Btn) then Btn.OnClick:=ClickARendre;
end;


procedure TOF_PGMULTISELECTION.ClickAPrendre(Sender: TObject);
begin
  ChangeGridObjet(GDResult,GdSelect,Tob_Result,Tob_Select);
end;

procedure TOF_PGMULTISELECTION.ClickARendre(Sender: TObject);
begin
  ChangeGridObjet(GdSelect,GDResult,Tob_Select,Tob_Result);
end;



procedure TOF_PGMULTISELECTION.ChangeGridObjet(Sender, Source: TObject; Tob_Sender, Tob_Source : Tob);
Var nbre,i : integer;
    TD,T   : Tob;
begin
  Nbre := THGrid(Source).RowCount - 1;
  if not THGrid(Source).MultiSelect then Nbre := 1;
  for i := 1 to Nbre do
  begin
    if THGrid(Source).IsSelected(i) then
    begin
      if not THGrid(Source).MultiSelect then
          TD := TOB(THGrid(Source).Objects[0, THGrid(Source).Row])
        else
          TD := TOB(THGrid(Source).Objects[0, i]);
        if not Assigned(TD) then exit;

        T := Tob_Source.FindFirst(['CODE'], [TD.GetValue('CODE')], FALSE);
        if T <> nil then T.ChangeParent(Tob_Sender, -1) else PgiBox('Pas trouvé','');
        //TD.Free;
    end;
  End;
  THGrid(Source).ClearSelected;
  AfficheGrid;
end;

procedure TOF_PGMULTISELECTION.AfficheGrid;
begin
  If (Assigned(Tob_Select))  then //and (Tob_Select.detail.count>0)
     if assigned(GDSelect) then
       Begin
       Tob_Select.PutGridDetail(GDSelect,False,False,'',True);
       GDSelect.SortGrid(0, FALSE);
       End;

  If (Assigned(Tob_Result)) then //and (Tob_Result.detail.count>0) then
     if assigned(GDResult) then
       Begin
       Tob_Result.PutGridDetail(GDResult,False,False,'',True);
       GDResult.SortGrid(0, FALSE);
       End;

end;


procedure TOF_PGMULTISELECTION.OnUpdate;
Var Tob_Update,TFille : Tob;
    i : Integer;
begin
  inherited;
ExecuteSql('DELETE FROM PGMULTISELECTION WHERE PMZ_MULTISELECTION="'+GetControlText('PMZ_MULTISELECTION')+'"');

If assigned(Tob_Result) and (Tob_Result.Detail.count > 0) then
  Begin
  Tob_Update := Tob.Create('PGMULTISELECTION',nil,-1);
  For i := 0 to Tob_Result.Detail.count-1 do
    Begin
    TFille := Tob.Create('PGMULTISELECTION',Tob_Update,-1);
    TFille.PutValue('PMZ_MULTISELECTION',GetControlText('PMZ_MULTISELECTION'));
    TFille.PutValue('PMZ_LIBELLE'       ,GetControlText('PMZ_LIBELLE'));
//    TFille.PutValue('PMZ_TABLE'         ,'SALARIES');//PrefixeToTable(Copy(Champ,1,3)));
    TFille.PutValue('PMZ_CHAMP'         ,Champ);
    TFille.PutValue('PMZ_VALEUR'        ,Tob_Result.Detail[i].GetValue('CODE'));
    End;
  Tob_Update.InsertOrUpdateDB;
  TFVierge(Ecran).Retour := GblTypeAction;
  End;


end;

procedure TOF_PGMULTISELECTION.OnClose;
begin
  inherited;
FreeAndNil(Tob_Select);
FreeAndNil(Tob_Result);
FreeAndNil(Tob_Delete);
AvertirTable('PGMULTISELECTION');
end;


procedure TOF_PGMULTISELECTION.OnDelete;
begin
  inherited;
ExecuteSql('DELETE FROM PGMULTISELECTION WHERE PMZ_MULTISELECTION="'+GetControlText('PMZ_MULTISELECTION')+'"');
TFVierge(Ecran).Retour := 'SUPPR';
end;

initialization
  registerclasses([TOF_PGMULTISELECTION]);

end.
