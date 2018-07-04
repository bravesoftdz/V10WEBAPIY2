unit UTOMPGMultitable;

interface

uses Sysutils,Classes,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,HEnt1,Utom,HTB97,HCtrls,Utob,Utof;

type
  TOM_PgMultiTable = class(TOM)
    procedure OnArgument(stArgument: string);    override;
    procedure OnLoadRecord;                      override;
    procedure OnNewRecord;                       override;
    procedure OnChangeField(F: TField);          override;
    procedure OnClose;                           override;
{    procedure OnUpdateRecord;                    override;
    procedure OnDeleteRecord;                    override;}
  {  procedure OnAfterUpdateRecord;               override;}
  private
    Tob_TableValeur : tob;
    Procedure ClickBtnTable(Sender : TObject);
    Procedure ClickBtnDetail (Sender : TObject);
   end;

   TOF_PgMultiTableMul = class(TOF)
   procedure OnArgument(stArgument: string);    override;
   End;



implementation
uses EntPaie; 

{ TOM_PgMultiTable }

procedure TOM_PgMultiTable.OnArgument(stArgument: string);
Var Btn  : TToolBarButton97;
    Q    : TQuery;
begin
  inherited;
Btn:=TToolBarButton97(GetControl('BTNTABLE1'));
If Assigned(Btn) then Btn.onclick := ClickBtnTable;
Btn:=TToolBarButton97(GetControl('BTNTABLE2'));
If Assigned(Btn) then Btn.onclick := ClickBtnTable;
Btn:=TToolBarButton97(GetControl('BTNTABLE3'));
If Assigned(Btn) then Btn.onclick := ClickBtnTable;
Btn:=TToolBarButton97(GetControl('BTNDETAIL'));
If Assigned(Btn) then Btn.onclick := ClickBtnDetail;
Q := OpenSQL('SELECT * FROM PGMULTITABLE WHERE ##PMU_PREDEFINI## PMU_TYPETABLE="VAL"',True);
If Not Q.eof then
  Begin
  Tob_TableValeur := Tob.Create('Les sélections',nil,-1);
  Tob_TableValeur.LoadDetailDB('PGMULTITABLE','','',Q,False);
  End;
Ferme(Q);
end;

procedure TOM_PgMultiTable.OnLoadRecord;
begin
  inherited;
end;


procedure TOM_PgMultiTable.OnNewRecord;
begin
  inherited;
Setfield('PMU_NODOSSIER','000000');
Setfield('PMU_TYPETABLE','TAB');
end;



procedure TOM_PgMultiTable.ClickBtnTable(Sender: TObject);
Var St,St2,Tablette,Sti : String;
    T               : Tob;
begin
If Not Assigned(Sender) then exit;
Sti := Copy(TToolBarButton97(Sender).Name,Length(TToolBarButton97(Sender).Name),1);
If not IsNumeric(Sti) then exit;
St2 := '';
Tablette := '';
St:='TABLEFILTRE'+Sti;
FreeAndNil(T);
if Assigned(Tob_TableValeur) then
   T := Tob_TableValeur.FindFirst(['PMU_MULTICODE'],[''+GetField('PMU_'+St)+''],False);
If Assigned(T) then
  Begin
  Tablette := T.GetValue('PMU_NOMTABLETTE');
  St2      := T.GetValue('PMU_NOMCHAMP');
  End;
St:='PMU_SELECTFILTRE'+Copy(TToolBarButton97(Sender).Name,Length(TToolBarButton97(Sender).Name),1);
St:= GetField(St);
St := AglLanceFiche('PAY','PGMULTISELECTION','','',st+';'+St2+';'+Tablette);
if St <> '' then
  Begin
  SetControlProperty('PMU_SELECTFILTRE'+Sti,'Plus','');
  If (Assigned(Tob_TableValeur)) then
     Begin
     T := Tob_TableValeur.FindFirst(['PMU_MULTICODE'],[''+GetField('PMU_TABLEFILTRE'+Sti)+''],False);
     If Assigned(T) then
       SetControlProperty('PMU_SELECTFILTRE'+Sti,'Plus','WHERE PMZ_CHAMP="'+T.GetValue('PMU_NOMCHAMP')+'" ');
     End;
  SetField('PMU_SELECTFILTRE'+Sti,'');
  SetFocusControl('PMU_TABLEFILTRE'+Sti);
  End;
end;


procedure TOM_PgMultiTable.OnChangeField(F: TField);
Var i : integer;
    Sti : String;
    T : Tob;
begin
  inherited;
  if (F.FieldName = 'PMU_NBFILTRE') then
    Begin
    For i:= 1 to 3 do
      Begin
      SetControlEnabled('PMU_TABLEFILTRE'+IntToStr(i),(GetField('PMU_NBFILTRE') >= i));
      if (GetField('PMU_NBFILTRE') < i) and (GetField('PMU_TABLEFILTRE'+IntToStr(i))<>'') then SetField('PMU_TABLEFILTRE'+IntToStr(i),'');
      End;
    //SetControlEnabled('PMU_TYPERESULTAT',(GetField('PMU_NBFILTRE') > 0));
    End;

  if (Pos('PMU_TABLEFILTRE',F.FieldName)>0) then
     Begin
     Sti := Copy(F.FieldName,Length(F.FieldName),1);
     If ISNumeric(Sti) then
       Begin
       { Rend accessible les zones liées }
       SetControlEnabled('PMU_SELECTFILTRE'+Sti,(GetField('PMU_TABLEFILTRE'+Sti)<>''));
       { Suppression valeur zone liée si aucun }
       if  (Ds.State in [DsInsert,DsEdit]) and (GetField('PMU_SELECTFILTRE'+Sti)<>'') then
               SetField('PMU_SELECTFILTRE'+Sti,'');
       { Applique .plus tablette zone liée }
       If (GetField('PMU_TABLEFILTRE'+Sti)<>'') AND (Assigned(Tob_TableValeur)) then
         Begin
         T := Tob_TableValeur.FindFirst(['PMU_MULTICODE'],[''+GetField('PMU_TABLEFILTRE'+Sti)+''],False);
         If Assigned(T) then SetControlProperty('PMU_SELECTFILTRE'+Sti,'Plus','WHERE PMZ_CHAMP="'+T.GetValue('PMU_NOMCHAMP')+'" ');
         End
       else
         SetControlProperty('PMU_SELECTFILTRE'+Sti,'Plus','');
       SetControlEnabled('BTNTABLE'+Sti,(GetField('PMU_TABLEFILTRE'+Sti)<>'') );
       End;
     End;

    if (F.FieldName = 'PMU_NBTABLE') then
    Begin
    For i:= 1 to 3 do
      Begin
      SetControlEnabled('PMU_TABLETRAVAIL'+IntToStr(i),(GetField('PMU_NBTABLE') >= i));
      if (GetField('PMU_NBTABLE') < i) and (GetField('PMU_TABLETRAVAIL'+IntToStr(i))<>'') then SetField('PMU_TABLETRAVAIL'+IntToStr(i),'');
      End;
    SetControlEnabled('PMU_NBVALEUR',(GetField('PMU_NBTABLE') > 0));   
    End;


    if (Pos('PMU_TABLETRAVAIL',F.FieldName)>0) then
     Begin
     Sti := Copy(F.FieldName,Length(F.FieldName),1);
     If ISNumeric(Sti) then
       Begin
       { Rend accessible les zones liées }
       SetControlEnabled('PMU_TABLECOND'+Sti,(GetField(F.FieldName)<>''));
       { Suppression valeur zone liée si aucun }
       if  (Ds.State in [DsInsert,DsEdit]) and (GetField('PMU_TABLECOND'+Sti)<>'') then
               SetField('PMU_TABLECOND'+Sti,'');
       End;
     End;


 { if (Pos('PMU_SELECTFILTRE',F.FieldName)>0) then
     Begin
     If ISNumeric(Copy(F.FieldName,Length(F.FieldName),1)) then
       SetControlEnabled('BTNTABLE'+Copy(F.FieldName,Length(F.FieldName),1),(GetField('PMU_SELECTFILTRE'+Copy(F.FieldName,Length(F.FieldName),1))<>'') );
     End;   }

  {if (F.FieldName = 'PMU_TYPERESULTAT') then
    Begin
    If Assigned(GdMultiDetail) then
    For i:=0 to GdMultiDetail.RowCount-1 do
      GdMultiDetail.Cells[GdMultiDetail.ColCount-1,i]:='';
    End;   }

end;





procedure TOM_PgMultiTable.ClickBtnDetail(Sender: TObject);
begin
AglLanceFiche('PAY', 'PGMULTIDETAIL', '', '', GetField('PMU_MULTICODE'));
end;

{ TOF_PgMultiTable_Mul }

procedure TOF_PgMultiTableMul.OnArgument(stArgument: string);
begin
  inherited;
If (VH_PAIE.PGLibelleOrgStat1 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PGLibelleOrgStat1+'" WHERE PMU_NOMCHAMP="PSA_TRAVAILN1"');
If (VH_PAIE.PGLibelleOrgStat2 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PGLibelleOrgStat2+'" WHERE PMU_NOMCHAMP="PSA_TRAVAILN2"');
If (VH_PAIE.PGLibelleOrgStat3 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PGLibelleOrgStat3+'" WHERE PMU_NOMCHAMP="PSA_TRAVAILN3"');
If (VH_PAIE.PGLibelleOrgStat4 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PGLibelleOrgStat4+'" WHERE PMU_NOMCHAMP="PSA_TRAVAILN4"');
If (VH_PAIE.PGLibCodeStat <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PGLibCodeStat+'" WHERE PMU_NOMCHAMP="PSA_CODESTAT"');
If (VH_PAIE.PgLibCoche1 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCoche1+'" WHERE PMU_NOMCHAMP="PSA_BOOLLIBRE1"');
If (VH_PAIE.PgLibCoche2 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCoche2+'" WHERE PMU_NOMCHAMP="PSA_BOOLLIBRE2"');
If (VH_PAIE.PgLibCoche3 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCoche3+'" WHERE PMU_NOMCHAMP="PSA_BOOLLIBRE3"');
If (VH_PAIE.PgLibCoche4 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCoche4+'" WHERE PMU_NOMCHAMP="PSA_BOOLLIBRE4"');
If (VH_PAIE.PgLibCombo1 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCombo1+'" WHERE PMU_NOMCHAMP="PSA_LIBREPCMB1"');
If (VH_PAIE.PgLibCombo2 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCombo2+'" WHERE PMU_NOMCHAMP="PSA_LIBREPCMB2"');
If (VH_PAIE.PgLibCombo3 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCombo3+'" WHERE PMU_NOMCHAMP="PSA_LIBREPCMB3"');
If (VH_PAIE.PgLibCombo4 <> '') then
   ExecuteSql('UPDATE PGMULTITABLE SET PMU_LIBELLE="'+VH_PAIE.PgLibCombo4+'" WHERE PMU_NOMCHAMP="PSA_LIBREPCMB4"');

(*
Lib := 'PGTABLETRAVAIL';
iTab := TTTONum(Lib) ;
if iTab > 0 then
   Begin
   if not Assigned(V_PGI.DECOMBOS[iTab].Valeurs) then exit;
   For i := 0 to V_PGI.DECOMBOS[iTab].Valeurs.Count-1 do
     Begin
     //'002'#9'Salarié'#9'PSA_SALARIE'#9'PGSALARIE'#9
     Lib:= V_PGI.DECOMBOS[iTab].Valeurs.Strings[i];
     If (Pos('PSA_TRAVAILN1',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PGLibelleOrgStat1 <> '') then Insert(VH_PAIE.PGLibelleOrgStat1,Lib,5)
       else                                      Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

     If (Pos('PSA_TRAVAILN2',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PGLibelleOrgStat2 <> '') then Insert(VH_PAIE.PGLibelleOrgStat2,Lib,5)
       else                                      Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

     If (Pos('PSA_TRAVAILN3',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PGLibelleOrgStat3 <> '') then Insert(VH_PAIE.PGLibelleOrgStat3,Lib,5)
       else                                      Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_TRAVAILN4',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PGLibelleOrgStat4 <> '') then Insert(VH_PAIE.PGLibelleOrgStat4,Lib,5)
       else                                      Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_CODESTAT',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PGLibCodeStat <> '') then Insert(VH_PAIE.PGLibCodeStat,Lib,5)
       else                                  Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_BOOLLIBRE1',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCoche1 <> '') then Insert(VH_PAIE.PgLibCoche1,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_BOOLLIBRE2',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCoche2 <> '') then Insert(VH_PAIE.PgLibCoche2,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_BOOLLIBRE3',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCoche3 <> '') then Insert(VH_PAIE.PgLibCoche3,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_BOOLLIBRE4',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCoche4 <> '') then Insert(VH_PAIE.PgLibCoche4,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_LIBREPCMB1',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCombo1 <> '') then Insert(VH_PAIE.PgLibCombo1,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_LIBREPCMB2',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCombo2 <> '') then Insert(VH_PAIE.PgLibCombo2,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_LIBREPCMB3',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCombo3 <> '') then Insert(VH_PAIE.PgLibCombo3,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;

      If (Pos('PSA_LIBREPCMB4',Lib) > 0) And (Pos('Non affecté',Lib) > 0) then
       Begin
       System.Delete(Lib,Pos('Non affecté',Lib),11);
       If (VH_PAIE.PgLibCombo4 <> '') then Insert(VH_PAIE.PgLibCombo4,Lib,5)
       else                                Insert('Non utilisé',Lib,5);
       V_PGI.DECOMBOS[iTab].Valeurs.Strings[i]:=Lib;
       Continue;
       End;
  {   Lib := VH_PAIE.PgLibCombo4;         if Trim(lib) = '' then Lib := 'Non utilisé';
   V_PGI.DECOMBOS[iTab].Valeurs.Strings['103']:= Lib;}
     End;
   V_PGI.DECOMBOS[iTab].Where := V_PGI.DECOMBOS[iTab].Where +' AND CO_LIBELLE<>"Non utilisé"';
   End;  *)
end;

procedure TOM_PgMultiTable.OnClose;
begin
  inherited;
FreeAndNil(Tob_TableValeur);
end;


initialization
  registerclasses([TOM_PgMultiTable,TOF_PgMultiTableMul]);
end.
