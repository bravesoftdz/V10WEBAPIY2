unit UtofPgMultiDetail;

interface

uses Classes,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}SysUtils,
     HCtrls,Utob,Htb97,Utof,Vierge;

type
  TOF_PGMULTIDETAIL = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    GdMultiDetail : THGrid;
    Tob_Grid,Tob_PgMultiTable,Tob_PgMultiValeur : Tob;
    Procedure ClickBtnGrille(Sender : TObject);
    Procedure ChangeCbTableTravail(Sender : TObject);
    Procedure ChargeTob_SelectFiltre(Champ,Sti : String; Var TTable : Tob);
    Function  PGGetTabTypeResultat : String;
    Function  PgGetIntitule(Champ,Valeur: String; AvecValeur : Boolean = False) : String;
    Function  CreateAndInitT_Grid (T : Tob; Lig : integer) : Tob;
    Procedure AfficheTitreColonne;
    Procedure RaffraichirGrille;
    Procedure ClickInsereObjet(Sender : TObject);
    Procedure ClickDeleteObjet(Sender : TObject);
  end;

implementation
uses PGOutils2;

{ TOF_PGMULTIDETAIL }

procedure TOF_PGMULTIDETAIL.OnArgument(Arguments: string);
Var
Btn : TToolBarButton97;
Q   : TQuery;
Cb  : THValComboBox;
i   : Integer;
T            : Tob;
begin
  inherited;

  Q := OpenSql('SELECT * FROM PGMULTITABLE WHERE PMU_MULTICODE="'+Arguments+'" AND PMU_TYPETABLE="TAB"',True);
  If Not Q.eof then
     Begin
     Tob_PgMultiTable := TOB.Create('PGMULTI_TABLE',nil,-1);
     Tob_PgMultiTable.LoadDetailDB('PGMULTI_TABLE','','',Q,False);
     Tob_PgMultiTable:=Tob_PgMultiTable.detail[0];
     End;
  Ferme(Q);

  Q := OpenSql('SELECT * FROM PGMULTITABLE WHERE PMU_TYPETABLE<>"TAB"',True);
  If Not Q.eof then
     Begin
     Tob_PgMultiValeur := TOB.Create('PGMULTI_TABLE',nil,-1);
     Tob_PgMultiValeur.LoadDetailDB('PGMULTI_TABLE','','',Q,False);
     End;
  Ferme(Q);



GdMultiDetail := THGrid(GetControl('GDMULTIDETAIL'));
Btn:=TToolBarButton97(GetControl('BTNGRILLE'));
If Assigned(Btn) then Btn.onclick := ClickBtnGrille;

Btn := TToolBarButton97 (GetControl('BTNINSERER'));
if Assigned(Btn) then Btn.OnClick := ClickInsereObjet;

Btn := TToolBarButton97 (GetControl('BTNDELETELINE'));
if Assigned(Btn) then Btn.OnClick := ClickDeleteObjet;




For i := 1 to 3 do
    Begin
    Cb:=THValComboBox(GetControl('CBTABLETRAVAIL'+IntToStr(i)));
    if Assigned(Cb) then Cb.Onclick:=ChangeCbTableTravail;
    If Assigned(Tob_PgMultiValeur) then
      Begin
      T := Tob_PgMultiValeur.FindFirst(['PMU_MULTICODE'],[Tob_PgMultiTable.GetValue('PMU_TABLEFILTRE'+IntToStr(i))],False);
      If Assigned(T) AND (T.GetValue('PMU_NOMTABLETTE')<>'') then
        SetControlProperty('CBTABLETRAVAIL'+IntToStr(i),'DataType',T.GetValue('PMU_NOMTABLETTE'));
      End;
    IF i > Tob_PgMultiTable.GetValue('PMU_NBFILTRE') then
      SetControlVisible('GB'+IntToStr(i),False);
    End;
//    SetControlProperty('CBTYPERESULTAT','DataType',PgGetTabTypeResultat);
end;

procedure TOF_PGMULTIDETAIL.OnClose;
begin
  inherited;
FreeAndNil(Tob_PgMultiTable);
FreeAndNil(Tob_PgMultiValeur);
end;

procedure TOF_PGMULTIDETAIL.OnUpdate;
begin
  inherited;

end;


procedure TOF_PGMULTIDETAIL.ClickBtnGrille(Sender: TObject);
Var i,j,IFiltre1,IFiltre2,IFiltre3,Iligne                   : integer;
    T,Tob_SelectFiltre1,Tob_SelectFiltre2,Tob_SelectFiltre3,TDouble : Tob;
    Filtre1,Filtre2 : Array [1..2] of String ;
begin
FreeAndNil(Tob_Grid);
If not assigned(GdMultiDetail) then exit;
GdMultiDetail.RowCount := 1;
GdMultiDetail.ColCount := 1;
if Tob_PgMultiTable.GetValue('PMU_NBFILTRE') + Tob_PgMultiTable.GetValue('PMU_NBTABLE') = 0 then Exit;
GdMultiDetail.ColCount := Tob_PgMultiTable.GetValue('PMU_NBFILTRE') + Tob_PgMultiTable.GetValue('PMU_NBTABLE') + 1;
GdMultiDetail.RowCount := 2 ;
GdMultiDetail.FixedRows := 1 ;
GdMultiDetail.FixedCols := Tob_PgMultiTable.GetValue('PMU_NBFILTRE');
{ Afffichege des titres de colonnes }
  AfficheTitreColonne;
{ Chargement des valeurs fixes de la grille }
  ChargeTob_SelectFiltre('PMU_SELECTFILTRE1','1',Tob_SelectFiltre1);
  ChargeTob_SelectFiltre('PMU_SELECTFILTRE2','2',Tob_SelectFiltre2);
  ChargeTob_SelectFiltre('PMU_SELECTFILTRE3','3',Tob_SelectFiltre3);

If Assigned(Tob_SelectFiltre1) And (Tob_SelectFiltre1.detail.count>0) then
  Begin
  Tob_Grid := Tob.Create('Les élements',Nil,-1);
  if PGGetTabTypeResultat<>'' then GdMultiDetail.colformats[GdMultiDetail.ColCount-1] := 'CB='+PGGetTabTypeResultat;
  {Balayage de la 1ère table de selection }
  Iligne := 0;
  For IFiltre1 := 0 to Tob_SelectFiltre1.detail.count-1 Do
     Begin
     Inc(Iligne);
     T:=CreateAndInitT_Grid(Tob_Grid,Iligne);
     Filtre1[1] := Tob_SelectFiltre1.detail[IFiltre1].GetValue('PMZ_VALEUR');
     Filtre1[2] := Tob_SelectFiltre1.detail[IFiltre1].GetValue('PMZ_CHAMP');
     T.PutValue('TABLETRAVAIL1',Filtre1[1]);
     T.PutValue('LIBTABLETRAVAIL1',PgGetIntitule(Filtre1[2],Filtre1[1],True));
     { Balayage de la 2ème table de selection }
     If Assigned(Tob_SelectFiltre2) And (Tob_SelectFiltre2.detail.count>0) then
       For IFiltre2 := 0 to Tob_SelectFiltre2.detail.count-1 Do
       Begin
       T.PutValue('TABLETRAVAIL1',Filtre1[1]);
       T.PutValue('LIBTABLETRAVAIL1',PgGetIntitule(Filtre1[2],Filtre1[1],True));
       Filtre2[1] := Tob_SelectFiltre2.detail[IFiltre2].GetValue('PMZ_VALEUR');
       Filtre2[2] := Tob_SelectFiltre2.detail[IFiltre2].GetValue('PMZ_CHAMP');
       T.PutValue('TABLETRAVAIL2',Filtre2[1]);
       T.PutValue('LIBTABLETRAVAIL2',PgGetIntitule(Filtre2[2],Filtre2[1],True));
       { Balayage de la 3ème table de selection }
       If Assigned(Tob_SelectFiltre3) And (Tob_SelectFiltre3.detail.count>0) then
         For IFiltre3 := 0 to Tob_SelectFiltre3.detail.count-1 Do
         Begin
         T.PutValue('TABLETRAVAIL1',Filtre1[1]);
         T.PutValue('LIBTABLETRAVAIL1',PgGetIntitule(Filtre1[2],Filtre1[1],True));
         T.PutValue('TABLETRAVAIL2',Filtre2[1]);
         T.PutValue('LIBTABLETRAVAIL2',PgGetIntitule(Filtre2[2],Filtre2[1],True));
         T.PutValue('TABLETRAVAIL3',Tob_SelectFiltre3.detail[IFiltre3].GetValue('PMZ_VALEUR'));
         T.PutValue('LIBTABLETRAVAIL3',PgGetIntitule(Tob_SelectFiltre3.detail[IFiltre3].GetValue('PMZ_CHAMP'),Tob_SelectFiltre3.detail[IFiltre3].GetValue('PMZ_VALEUR'),True));
         if (IFiltre3<Tob_SelectFiltre3.detail.count-1) then
           Begin
           Inc(Iligne);
           T:=CreateAndInitT_Grid(Tob_Grid,Iligne);
           End;
         End;
       if (IFiltre2<Tob_SelectFiltre2.detail.count-1) then
         Begin
         Inc(Iligne);
         T:=CreateAndInitT_Grid(Tob_Grid,Iligne);
         End;
       End;
     End;

  { Dedoublenage des lignes }
  if Tob_PgMultiTable.GetValue('PMU_NBVALEUR') > 1 then
    For j:=0 to  Tob_Grid.detail.count-1 Do
     Begin
     T := Tob_Grid.detail[j];
     For i := 1 to Tob_PgMultiTable.GetValue('PMU_NBVALEUR')-1 do
        Begin
        TDouble:=CreateAndInitT_Grid(Tob_Grid,Iligne+j);
        TDouble.Dupliquer(T,True,True);
        End;
     End;
  { Tri Tob + affichage de la grille }
  RaffraichirGrille;
  End;

//FreeAndNil(Tob_Grid);
FreeAndNil(Tob_SelectFiltre1);
FreeAndNil(Tob_SelectFiltre2);
FreeAndNil(Tob_SelectFiltre3);
end;


procedure TOF_PGMULTIDETAIL.ChargeTob_SelectFiltre(Champ,Sti : String; var TTable: Tob);
Var
T : Tob;
Q : TQuery;
begin
FreeAndNil(TTable);
if Tob_PgMultiTable.GetValue('PMU_TABLEFILTRE'+Sti) = '' then exit;
TTable := TOB.Create('PGMULTI_SELECTION1',nil,-1);
If Tob_PgMultiTable.GetValue(Champ)<>'' then
Begin
Q := OpenSql('SELECT * FROM PGMULTISELECTION WHERE PMZ_MULTISELECTION = "'+Tob_PgMultiTable.GetValue(Champ)+'"',True);
TTable.LoadDetailDB('PGMULTISELECTION','','',Q,False);
Ferme(Q);
End
else
  Begin
  T :=TOB.Create('PGMULTI_SELECTION1',TTable,-1);
  T.AddChampSupValeur('PMZ_VALEUR','Tous');
  End;
end;

function TOF_PGMULTIDETAIL.PgGetTabTypeResultat: String;
Var
  St : String;
begin
  Result := '';
  St := Tob_PgMultiTable.GetValue('PMU_TYPERESULTAT');
  if ((St = 'ANC') or (St = 'AGE')) then Result := 'PGTABINT'+St
  else if St = 'VAR' then Result := 'PGVARIABLE'
  else if St = 'ELT' then Result := 'PGELEMENTNAT';

end;

procedure TOF_PGMULTIDETAIL.ChangeCbTableTravail(Sender: TObject);
Var i : Integer;
    Affecter : Boolean;
begin
SetControlEnabled('BTNINSERER',False);
Affecter := True;
For i := 1 to Tob_PgMultiTable.GetValue('PMU_NBFILTRE') do
   if GetControlText('CBTABLETRAVAIL'+IntToStr(i)) = '' then
     Begin
     Affecter := False;
     Break;
     End;
SetControlEnabled('BTNINSERER',(Affecter=True));
end;

function TOF_PGMULTIDETAIL.PgGetIntitule(Champ,Valeur: String; AvecValeur : Boolean = False): String;
Var
   T : Tob;
begin
  Result := Valeur ;//' : '+RechDom(ArrayTablette[Col],Valeur,False);
  If Assigned(Tob_PgMultiValeur) then
    Begin
    T := Tob_PgMultiValeur.FindFirst(['PMU_NOMCHAMP'],[Champ],False);
    If Assigned(T) AND (T.GetValue('PMU_NOMTABLETTE')<>'') then
      If AvecValeur then Result := Valeur + ' : '+RechDom(T.GetValue('PMU_NOMTABLETTE'),Valeur,False)
      else Result := RechDom(T.GetValue('PMU_NOMTABLETTE'),Valeur,False);
    End;
end;

Function TOF_PGMULTIDETAIL.CreateAndInitT_Grid(T : Tob; Lig : integer) : Tob;
var i : integer;
begin
result := nil;
if Not Assigned (T) then exit;
Result:=Tob.create('les filles',T,-1);
Result.AddChampSupValeur('ORDRE',lig);
For i := 1 to Tob_PgMultiTable.GetValue('PMU_NBFILTRE') do
  Begin
  Result.AddChampSup('TABLETRAVAIL'+IntToStr(i),True);
  Result.AddChampSup('LIBTABLETRAVAIL'+IntToStr(i),True);
  End;
Result.AddChampSup('RESULTAT',True);
end;

procedure TOF_PGMULTIDETAIL.AfficheTitreColonne;
Var i : integer;
    St : String;
    T  : Tob;
begin
  GdMultiDetail.Titres.Clear;
  For i:= 1 to Tob_PgMultiTable.GetValue('PMU_NBFILTRE') do
    GdMultiDetail.Titres.Add(RechDom('PGTABLETRAVAIL',Tob_PgMultiTable.GetValue('PMU_TABLEFILTRE'+IntToStr(i)),False));
  For i:= 1 to Tob_PgMultiTable.GetValue('PMU_NBTABLE') do
    Begin
    St := RechDom('PGTABLETRAVAIL', Tob_PgMultiTable.GetValue('PMU_TABLETRAVAIL'+IntToStr(i)), FALSE);
    If Assigned(Tob_PgMultiValeur) then
      Begin
      T := Tob_PgMultiValeur.FindFirst(['PMU_MULTICODE'],[Tob_PgMultiTable.GetValue('PMU_TABLETRAVAIL'+IntToStr(i))],False);
      If Assigned(T) AND (T.GetValue('PMU_NOMTABLETTE')<>'') then
        St := St + ' '+ RechDom(T.GetValue('PMU_NOMTABLETTE'), Tob_PgMultiTable.GetValue('PMU_TABLECOND'+IntToStr(i)), FALSE);
      End;
    GdMultiDetail.Titres.Add(St);
    End;
  GdMultiDetail.Titres.Add(RechDom('PGTYPERESULTAT', Tob_PgMultiTable.GetValue('PMU_TYPERESULTAT'), FALSE));
  GdMultiDetail.UpdateTitres;
end;


procedure TOF_PGMULTIDETAIL.RaffraichirGrille;
Var i : integer;
    St,StLib : String;
begin
  if not Assigned(Tob_Grid) then exit;
  For i := 1 to Tob_PgMultiTable.GetValue('PMU_NBFILTRE') do
    Begin
    St := St + 'TABLETRAVAIL'+IntToStr(i)+';';
    StLib := StLib + 'LIBTABLETRAVAIL'+IntToStr(i)+';';
    End;
  Tob_Grid.detail.Sort(St);
  Tob_Grid.PutGridDetail(GdMultiDetail,False,False,StLib+'RESULTAT',True);
end;

procedure TOF_PGMULTIDETAIL.ClickInsereObjet(Sender: TObject);
Var T,T1 : Tob;
    i  : integer;
begin
If GdMultiDetail.RowCount = 0 then exit;
T:=CreateAndInitT_Grid(Tob_Grid,-1);
if Assigned(T) then
   Begin
   For i := 1 to Tob_PgMultiTable.GetValue('PMU_NBFILTRE') do
      Begin
      T.PutValue('TABLETRAVAIL'+IntToStr(i),GetControlText('CBTABLETRAVAIL'+IntToStr(i)));
      If Assigned(Tob_PgMultiValeur) then
        Begin
        T1 := Tob_PgMultiValeur.FindFirst(['PMU_MULTICODE'],[Tob_PgMultiTable.GetValue('PMU_TABLEFILTRE'+IntToStr(i))],False);
        If Assigned(T1) AND (T1.GetValue('PMU_NOMCHAMP')<>'') then
           T.PutValue('LIBTABLETRAVAIL'+IntToStr(i),PgGetIntitule(T1.GetValue('PMU_NOMCHAMP'),GetControlText('CBTABLETRAVAIL'+IntToStr(i)),True));
        End;
      End;
   RaffraichirGrille;
   End;
end;



procedure TOF_PGMULTIDETAIL.ClickDeleteObjet(Sender: TObject);
var ARow : integer;
begin
If GdMultiDetail.RowCount = 0 then exit
else
  if GdMultiDetail.RowCount>1 then
    Begin
    ARow := GdMultiDetail.Row;
    GdMultiDetail.CacheEdit;
    GdMultiDetail.SynEnabled := False;
    GdMultiDetail.DeleteRow(GdMultiDetail.Row);
    if Arow < 2 then GdMultiDetail.Row := 1 else GdMultiDetail.Row := ARow-1;
    GdMultiDetail.MontreEdit;
    GdMultiDetail.SynEnabled := True;
    GdMultiDetail.Col := GdMultiDetail.FixedCols;
    End;
end;

initialization
  registerclasses([TOF_PGMULTIDETAIL]);

end.
